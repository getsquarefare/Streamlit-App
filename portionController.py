import copy
import json
import re
import os
import sys
from tqdm import tqdm
import pandas as pd
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

# Local imports
from store_access import new_database_access
from DishOptimizerLLM import LLMDishOptimizer


if __name__ == "__main__":
    import os
    import sys

    """The following three lines may not be needed for an application context, 
    but are needed for direct main script execution"""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.abspath(os.path.join(current_dir, ".."))
    sys.path.append(project_dir)

class MealRecommendation:
    def __init__(self) -> None:
        # Initialize database connection
        self.db = new_database_access()  # default_store_access: testing database
        # Clear previous recommendation results
        # self.clear_previous_results()
        # Initialize recommendation ID to track each recommendation
        self.recommendation_id = 1

    def clear_previous_results(self):
        # Delete previous results
        # self.db.delete_all_client_error()
        self.db.delete_all_clientservings()

    def convert_to_nutrient_constraints(self, fields):
        # Configuration for mapping record fields to nutrient constraints
        field_mapping = {
            "Protein (g)": [("Protein LB", "lb"), ("Protein UB", "ub")],
            "Kcal": [("KCal LB", "lb"), ("KCal UB", "ub")],
            "Fat, Total (g)": [("Fat LB", "lb"), ("Fat UB", "ub")],
            "Dietary Fiber (g)": [("Fiber LB", "lb"), ("Fiber UB", "ub")],
            "Carbohydrate, total (g)": [("Carbs LB", "lb"), ("Carbs UB", "ub")],
        }

        # Initialize the constraints dictionary with default values
        nutrient_constraints = {}
        default_constraints = {"lb": 0.8, "ub": 1.2}

        # Loop through the field mapping to construct nutrient_constraints
        for nutrient, mappings in field_mapping.items():
            # Initialize the nutrient dictionary if it doesn't exist
            if nutrient not in nutrient_constraints:
                nutrient_constraints[nutrient] = {}
            for record_field, constraint_key in mappings:
                if record_field in fields:
                    # Add the specific constraint from the record fields
                    nutrient_constraints[nutrient][constraint_key] = fields[
                        record_field
                    ]
                else:
                    # Set to default value if not found in the record fields
                    nutrient_constraints[nutrient][constraint_key] = (
                        default_constraints[constraint_key]
                    )
        return nutrient_constraints
    
    def clean_up_dish(self, data):
        # Extract the dish name from the first record (assuming all records belong to the same dish)
        dish_name = data[0]['Airtable Dish Name']
        
        # Create a cleaned ingredients list
        ingredients = []
        for item in data:
            ingredients.append({
                "component": item['Component (from Ingredient)'][0],
                "ingredientName": item['Ingredient Name'].strip(),
                "grams": item['Grams'],
                "kcal": item['Kcal'],
                "protein(g)": item['Protein (g)'],
                "fat(g)": item['Fat, Total (g)'],
                "dietaryFiber(g)": item['Dietary Fiber (g)'],
                "carbohydrate(g)": item['Carbohydrate, total (g)']
            })
        
        # Return the final JSON structure
        return {
            "dishName": dish_name,
            "ingredients": ingredients
        }

    def optimize(self, dish, customer_requirements):
        # Define the list of nutritional variables to optimize
        nutrients = [
            "Kcal",
            "Carbohydrate, total (g)",
            "Protein (g)",
            "Fat, Total (g)",
            "Dietary Fiber (g)",
            "Grams",
        ]

        # Each customer has exactly one constraint - choose from existing ones or customize
        constraint_id = customer_requirements["Portion Algo Constraints"][0]
        constraints: dict = self.db.get_constraints_details_by_rcdId(id=constraint_id)
        constraints: dict = constraints.get("fields", {})

        # Retrieve nutrient constraints
        # Param - Nutrient Constraints: KCal/Carb/Protein/Fat/Fiber UB/LB
        nutrient_constraints = self.convert_to_nutrient_constraints(constraints)

        # Param - Customer Requirements: KCal/Carb/Protein/Fat/Fiber
        if "goal_calories" in customer_requirements:
            customer_requirements["Kcal"] = customer_requirements["goal_calories"]
        if "goal_carbs(g)" in customer_requirements:
            customer_requirements["Carbohydrate, total (g)"] = customer_requirements[
                "goal_carbs(g)"
            ]
        if "goal_protein(g)" in customer_requirements:
            customer_requirements["Protein (g)"] = customer_requirements[
                "goal_protein(g)"
            ]
        if "goal_fat(g)" in customer_requirements:
            customer_requirements["Fat, Total (g)"] = customer_requirements[
                "goal_fat(g)"
            ]
        if "goal_fiber(g)" in customer_requirements:
            customer_requirements["Dietary Fiber (g)"] = customer_requirements[
                "goal_fiber(g)"
            ]

        # Retrieve additional constraints
        # Param - Sauce Grams
        sauce_grams = constraints.get("Sauce Amount", None)
        # Param - Veggie Ge Starch
        veggie_ge_starch = constraints.get("Veggie >= Starch", None)
        # Param - Min Meat per 100 Cal
        min_meat_per_100_cal = constraints.get("Minimum Meat Per 100KCal", None)
        # Param - Max Meal Grams per 100 Cal
        max_meal_grams_per_100_cal = constraints.get(
            "Maximum Meal Grams Per 100KCal", None
        )

        garnish_grams = 0
        # Param - Grouped Ingredients: use a normal dictionary to group different ingredients by
        # their components to ensure that ingredients from a component are scaled together
        grouped_ingredients = {}

        for ingredient in dish:
            # Replace meat with protein
            ingredient["Component (from Ingredient)"][0] = ingredient["Component (from Ingredient)"][0].lower()
            if ingredient["Component (from Ingredient)"][0] == "meat":
                ingredient["Component (from Ingredient)"][0] = "protein"
            if ingredient["Component (from Ingredient)"][0] == "garnish":
                garnish_grams += ingredient["Grams"]
            
            component = ingredient["Component (from Ingredient)"][0]
            if component not in grouped_ingredients:
                grouped_ingredients[component] = {nutrient: 0.0 for nutrient in nutrients}
            
            for nutrient in nutrients:
                grouped_ingredients[component][nutrient] += ingredient[nutrient]

        dish = self.clean_up_dish(dish)

        optimizer = LLMDishOptimizer(
            grouped_ingredients,
            customer_requirements,
            nutrients,
            nutrient_constraints,
            garnish_grams,
            sauce_grams,
            veggie_ge_starch,
            min_meat_per_100_cal,
            max_meal_grams_per_100_cal,
            dish,
        )

        try:
            response = optimizer.solve()
            return response
        except:
            return None

    # Function to aggregate grams by component
    def summarize_components(self, dish):
        components = {}
        for ingredient in dish:
            comp = ingredient["Component (from Ingredient)"][0]
            name = ingredient["Ingredient Name"]
            grams = ingredient["Grams"]
            if comp not in components:
                components[comp] = {}
            if name in components[comp]:
                components[comp][name] += grams
            else:
                components[comp][name] = grams
        return components

    # Create the dict to fill the cells in table recommendation summary later
    def get_recommendation_summary(
        self,
        dish_name,
        recommendation_id,
        shopify_id,
        dish,
        client,
        nutritional_information,
        final_ingredients,
        deletions,
        explanation,
        review_needed
    ):
        meat_g = sum(
            [
                ingredient["Grams"]
                for ingredient in dish
                if ingredient["Component"] == "protein"
            ]
        )
        veggies_g = sum(
            [
                ingredient["Grams"]
                for ingredient in dish
                if ingredient["Component"] == "veggies"
            ]
        )
        sauce_g = sum(
            [
                ingredient["Grams"]
                for ingredient in dish
                if ingredient["Component"] == "sauce"
            ]
        )
        starch_g = sum(
            [
                ingredient["Grams"]
                for ingredient in dish
                if ingredient["Component"] == "starch"
            ]
        )

        garnish_g = sum(
            [
                ingredient["Grams"]
                for ingredient in dish
                if ingredient["Component"] == "garnish"
            ]
        )

        recommendation_summary = {
            "Recommendation ID": recommendation_id,
            "Linked OrderItem": self.db.get_rcdid_by_shopify_orderlineitem(shopify_id),
            "Dish": dish_name,
            "Customer_FName": client["First_Name"],
            "Customer_LName": client["Last_Name"],
            "Delivery Date": "NA",
            "All Deletions": deletions,
            "Meat": self.get_ingrdts_one_component("Meat", final_ingredients),
            "Meat (g)": meat_g,
            "Sauce": self.get_ingrdts_one_component("Sauce", final_ingredients),
            "Sauce (g)": sauce_g,
            "Starch": self.get_ingrdts_one_component("Starch", final_ingredients),
            "Starch (g)": starch_g,
            "Veggies": self.get_ingrdts_one_component("Veggies", final_ingredients),
            "Veggies (g)": veggies_g,
            "Garnish": self.get_ingrdts_one_component("Garnish", final_ingredients),
            "Garnish (g)": garnish_g,
            "Total Calories (kcal)": nutritional_information["Calories"],
            "Total Carbs (g)": nutritional_information["Carbohydrates"],
            "Total Protein (g)": nutritional_information["Protein"],
            "Total Fat (g)": nutritional_information["Fat"],
            "Total Fiber (g)": nutritional_information["Fiber"],
            "Updated Nutrition Info": str(nutritional_information),
            "Review Needed": review_needed,
            "Explanation": explanation
        }
        return recommendation_summary

    def get_default_recommendation_summary(
        self, dish_name, shopify_id, dish, client, final_ingredients, deletions,explanation
    ):
        meat_g = sum(
            [
                ingredient["Grams"]
                for ingredient in dish
                if ingredient["Component (from Ingredient)"][0] == "protein"
            ]
        )
        veggies_g = sum(
            [
                ingredient["Grams"]
                for ingredient in dish
                if ingredient["Component (from Ingredient)"][0] == "veggies"
            ]
        )
        sauce_g = sum(
            [
                ingredient["Grams"]
                for ingredient in dish
                if ingredient["Component (from Ingredient)"][0] == "sauce"
            ]
        )
        starch_g = sum(
            [
                ingredient["Grams"]
                for ingredient in dish
                if ingredient["Component (from Ingredient)"][0] == "starch"
            ]
        )

        garnish_g = sum(
            [
                ingredient["Grams"]
                for ingredient in dish
                if ingredient["Component (from Ingredient)"][0] == "garnish"
            ]
        )

        recommendation_summary = {
            "Recommendation ID": dish[0]["Recommendation ID"],
            "Linked OrderItem": self.db.get_rcdid_by_shopify_orderlineitem(shopify_id),
            "Dish": dish_name,
            "Customer_FName": client["First_Name"],
            "Customer_LName": client["Last_Name"],
            "Delivery Date": "NA",
            #"All Deletions": deletions,
            "Meat": self.get_ingrdts_one_component("Meat", final_ingredients),
            "Meat (g)": meat_g,
            "Sauce": self.get_ingrdts_one_component("Sauce", final_ingredients),
            "Sauce (g)": sauce_g,
            "Starch": self.get_ingrdts_one_component("Starch", final_ingredients),
            "Starch (g)": starch_g,
            "Veggies": self.get_ingrdts_one_component("Veggies", final_ingredients),
            "Veggies (g)": veggies_g,
            "Garnish": self.get_ingrdts_one_component("Garnish", final_ingredients),
            "Garnish (g)": garnish_g,
            "Updated Nutrition Info": "Using default recipe",
            "Review Needed": True,
            "Explanation": explanation,
        }
        return recommendation_summary

    def get_ingrdts_one_component(self, component, final_ingredients):
        results = []
        for final_ingredient in final_ingredients:
            ingredient = self.db.get_ingredient_details_by_recId(final_ingredient)
            if ingredient["Component"] == component:
                results.append(ingredient["Ingredient Name"])
        return results

    def get_component_info(self, component, final_ingredients, deletions):
        contained = set()
        deleted = set()
        for final_ingredient in final_ingredients:
            ingredient = self.db.get_ingredient_details_by_recId(final_ingredient)
            if ingredient["Component"] == component:
                contained.add(ingredient["Ingredient Name"])
        for deletion in deletions:
            ingredient = self.db.get_ingredient_details_by_recId(deletion)
            if ingredient["Component"] == component:
                deleted.add(ingredient["Ingredient Name"])
        message = "Contains: "
        for ingredient in contained:
            message = message + ingredient + ", "
        if len(deleted) != 0:
            message += "\n Deletions: "
        for ingredient in deleted:
            message = message + ingredient + ", "
        return message

    def get_dish_nutritional_information(self, dish):
        """
        Calculate the total nutritional content for a specific dish and ingredient type.

        Parameters:
        - dish_name (list): List of ingredient specifications that make up the dish.

        Returns:
        - dict: A dictionary containing the total grams, calories, carbs, protein, fat, and fiber.
        """

        # Sum up the nutritional content
        dish_df = pd.DataFrame(dish)

        # Sum up the nutritional content
        total_grams = int(dish_df["Grams"].sum())
        total_calories = int(dish_df["Kcal"].sum())
        total_carbs = int(dish_df["Carbohydrate, total (g)"].sum())
        total_protein = int(dish_df["Protein (g)"].sum())
        total_fat = int(dish_df["Fat, Total (g)"].sum())
        total_fiber = int(
            dish_df["Dietary Fiber (g)"].fillna(0).sum()
        )  # Handle NaN values for fiber

        return {
            "Total Grams": total_grams,
            "Total Calories": total_calories,
            "Total Carbs (g)": total_carbs,
            "Total Protein (g)": total_protein,
            "Total Fat (g)": total_fat,
            "Total Fiber (g)": total_fiber,
        }

    def generate_recommendations(self):
        # self.db.delete_all_clientservings() # only when reset
        open_orders = self.db.get_all_open_orders()
        client_dish_pairs = self.build_client_dish_mapping(
            open_orders,
            shopify_id_column="SquareSpace/ Internal OrderItem ID",
            client_column="To_Match_Client_Nutrition",
            dish_column="Dish ID",
            ingredient_column="Final Ingredients",
            deletion_column="Deletions",
        )

        # Create a ThreadPoolExecutor to run tasks concurrently
        with ThreadPoolExecutor(max_workers=3) as executor:
            future_to_pair = {
                executor.submit(self.process_recommendation, shopify_id, client_id, dish_id, final_ingredients, deletions): 
                (shopify_id, client_id, dish_id) 
                for shopify_id, client_id, dish_id, final_ingredients, deletions in client_dish_pairs
            }

            for future in as_completed(future_to_pair):
                shopify_id, client_id, dish_id = future_to_pair[future]
                try:
                    future.result()  # Retrieve result or raise exception if occurred
                except Exception as e:
                    print(f"Error processing recommendation for Dish ID {dish_id} (Client ID {client_id}): {e}")

    def process_recommendation(self, shopify_id, client_id, dish_id, final_ingredients, deletions):
        # Extracted recommendation logic for concurrent execution in generate_recommendations
        try:
            client = self.db.get_client_details(recId=client_id)
            dish = self.db.get_dish_calc_nutritions_by_dishId(dish_id=dish_id)
        except Exception as e:
            print(f"Missing information for Dish ID {dish_id} (Client ID {client_id}): {e}")
            return  # Exit this task if information is missing

        if not client or not dish:
            print(f"Client or dish data missing for Dish ID {dish_id} (Client ID {client_id}).")
            return

        final_ingredients_set = set()
        orig_ingredients_set = set()
        final_dish = []

        for final_ingredient in final_ingredients:
            try:
                final_ingredients_set.add(
                    self.db.get_ingredient_details_by_recId(final_ingredient)["Ingredient ID"]
                )
            except Exception as e:
                print(f"Error fetching ingredient {final_ingredient} details: {e}")

        # Populate original and final dish details
        for ingredient in dish:
            orig_ingredients_set.add(ingredient["Ingredient ID"])
            if ingredient["Ingredient ID"] in final_ingredients_set:
                ingredient["Recommendation ID"] = self.recommendation_id
                ingredient["Dish ID"] = dish_id
                ingredient["Client_Id"] = [client["identifier"]]
                final_dish.append(ingredient)

        # Add new ingredients not present in the original recipe
        for ingredient_recId in final_ingredients:
            ingredient = self.db.get_ingredient_details_by_recId(ingredient_recId)
            if ingredient["Ingredient ID"] not in orig_ingredients_set:
                ingredient["Component (from Ingredient)"] = [ingredient["Component"]]
                ingredient["Kcal"] = ingredient["Energy (kcal)"]
                component = ingredient["Component"]
                ingredient["Grams"] = 5 if component == "Garnish" else 20 if component == "Sauce" else 50 if component == "Veggies" else 200
                final_dish.append(ingredient)

        if not final_dish:
            return

        dish_name = final_dish[0]["Airtable Dish Name"]
        recommendation_id = final_dish[0]["Recommendation ID"]

        try:
            # Run optimization and process response
            response = self.optimize(final_dish, client)
            # Use regex to extract the JSON part from the response as the modified recipe
            json_part = json.loads(re.search(r"\{.*\}", response, re.DOTALL).group(0))
            recipe = json_part.get("modified_recipe", {}).get("ingredients", [])
            nutritional_information = json_part.get(
                "updated_nutrition_info", {}
            )
            explanation = json_part.get("explanation", {})
            review_needed = bool(json_part.get("review_needed"))
            recommendation_summary = self.get_recommendation_summary(
                dish_name,
                recommendation_id,
                shopify_id,
                recipe,
                client,
                nutritional_information,
                final_ingredients=final_ingredients,
                deletions=deletions,
                explanation=explanation,
                review_needed=review_needed
            )
            self.db.output_clientservings(recommendation_summary)
        except Exception as e:
            # Fallback to default recommendation summary in case of failure
            print(f"Error in optimization for {shopify_id} (Client ID {client_id}): {e}")
            #  print("Original response: " + str(response))
            default_recommendation_summary = self.get_default_recommendation_summary(
                dish_name,
                shopify_id,
                final_dish,
                client,
                final_ingredients,
                deletions,
                explanation = 'Error in optimization',
            )
            self.db.output_clientservings(default_recommendation_summary)

    def build_client_dish_mapping(
        self,
        open_orders,
        shopify_id_column,
        client_column,
        dish_column,
        ingredient_column,
        deletion_column,
    ):
        client_dish_pairs = []

        for open_order in open_orders:
            if "fields" not in open_order:
                continue
            record_data = open_order["fields"]

            if client_column not in record_data or dish_column not in record_data:
                continue
            final_ingredients = []
            deletions = []

            if ingredient_column is not None and ingredient_column in record_data:
                final_ingredients = record_data[ingredient_column]
            if deletion_column is not None and deletion_column in record_data:
                deletions = record_data[deletion_column]
            shopify_id = record_data[shopify_id_column]
            client_id = record_data[client_column][0]
            dish_id = record_data[dish_column]

            client_dish_pairs.append(
                (shopify_id, client_id, dish_id, final_ingredients, deletions)
            )

        return client_dish_pairs


if __name__ == "__main__":
    meal_recommendation = MealRecommendation()
    meal_recommendation.generate_recommendations()
