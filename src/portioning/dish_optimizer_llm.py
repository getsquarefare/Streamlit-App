from openai import OpenAI
import os
from dotenv import load_dotenv
import streamlit as st

MAX_VEGGIES_GRAM = 300
class LLMDishOptimizer:
    def __init__(self, grouped_ingredients, customer_requirements, nutrients, nutrient_constraints, garnish_grams=None, sauce_grams=None, veggie_ge_starch=None, min_meat_per_100_cal=None, max_meal_grams_per_100_cal=None,dish=None):
        self.grouped_ingredients = grouped_ingredients
        self.customer_requirements = customer_requirements
        self.nutrients = nutrients
        self.nutrient_constraints = nutrient_constraints
        self.garnish_grams = garnish_grams
        self.sauce_grams = sauce_grams
        self.veggie_ge_starch = veggie_ge_starch
        self.min_meat_per_100_cal = min_meat_per_100_cal
        self.max_meal_grams_per_100_cal = max_meal_grams_per_100_cal
        self.dish = dish 

    def generate_prompt(self):
        # Base description
        base_description = (
        "You are an AI assistant tasked with modifying a recipe to meet specific customer nutrition goals. "
        "You will receive the customer's nutrition goals, a default recipe, and its nutritional information. "
        "Your task is to adjust the recipe's ingredient quantities (grams) to align as closely as possible with the customer's goals, "
        "while maintaining the dish's core essence.\n\n"
        )

        # Inputs and constraints
        inputs_constraints = (
        "Inputs:\n"
        f"Default Recipe: {self.dish}\n"
        # f"Default nutritional information: {self.grouped_ingredients}\n"
        f"Customer Requirements: {self.customer_requirements}\n"

        "Constraints which you must strictly follow:\n"
        f"The max total grams of veggies (ingredients in which components are 'veggies') are {MAX_VEGGIES_GRAM}\n"
        )

        # Optional fields
        optional_fields = ""
        if self.garnish_grams:
            optional_fields += f"Fixed Garnish Grams if any: {self.garnish_grams}\n"
        if self.sauce_grams:
            optional_fields += f"Fixed Sauce Grams if any: {self.sauce_grams}\n"
        if self.veggie_ge_starch:
            optional_fields += f"Veggie Grams need to be >= Starch Grams: {self.veggie_ge_starch}\n"
        if self.min_meat_per_100_cal:
            optional_fields += f"Min Meat Grams per 100 Cal: {self.min_meat_per_100_cal}\n"
        if self.max_meal_grams_per_100_cal:
            optional_fields += f"Max Meal Grams per 100 Cal: {self.max_meal_grams_per_100_cal}\n"

        # Explanation of inputs
        explanation_of_inputs = (
        "Explanation of Inputs and Constraints:\n"
        "- **Grouped Ingredients:** A breakdown of the ingredients grouped by their components such as veggies, starch, protein, etc., along with their nutritional profiles.\n"
        "- **Customer Requirements:** The customer's specific nutrition goals, including calorie, macronutrient, and fiber targets, as well as other constraints.\n"
        "- **lb and ub:**" 
        "  These are provided as a lower bound (lb) and an upper bound (ub) of each macronutrition relative to the customer's goal."
        "  For example, if a customer has a goal of 100g protein and the constraint is {'lb': 0.8, 'ub': 1.2}, the acceptable range for protein is 80g-120g in the final recipe."
        "  Similarly, if the calorie goal is 540 and the constraint is {'lb': 0.8, 'ub': 1.2}, the acceptable calorie range is 432-648.\n"
        "  Nutritional values within these ranges are considered valid and meet the requirements.\n\n"
        "- **Fixed Garnish Grams:** If any garnish grams are fixed and should not be adjusted.\n"
        "- **Fixed Sauce Grams:** If any sauce grams are fixed and should not be adjusted.\n"
        "- **Veggie Grams >= Starch Grams:** A boolean indicating whether the total grams of veggies should be greater than or equal to the total grams of starch in the final recipe.\n"
        "- **Min Meat Grams per 100 Cal:** The minimum grams of protein per 100 calories in the final recipe.\n"
        "- **Max Meal Grams per 100 Cal:** The maximum grams of the entire meal per 100 calories.\n\n"
        )

        # Step-by-step instructions
        step_by_step = (
        "Step-by-Step Instructions:\n\n"
        "1. Analyze the Nutritional Gap:\n"
        "   - Compare the nutritional profile of the default recipe to the customer's nutrition goals.\n"
        "   - Identify which nutrients need to be increased or decreased to meet the goals.\n\n"
        "2. Modify Ingredient Quantities:\n"
        "   - Adjust ingredient grams to balance nutrient levels using scaling.\n"
        "   - The grams don't need to be perfect like 200 or 150; values like 158 or 203 are acceptable.\n\n"
        "   - Try to achieve the protein and calories to be as closed to goals as you can first, then adjust other nutrients.\n"
        "3. Recalculate Nutritional Information:\n"
        "   - Recalculate the updated nutritional values of the modified recipe.\n\n"
        "4. Present the Modified Recipe:\n"
        "   - List the updated ingredient quantities and nutritional information in the following format:\n\n"
        "   <modified_recipe>\n"
        "   Ingredients:\n"
        "   [List of modified ingredients with updated quantities]\n\n"
        "   Updated Nutritional Information:\n"
        "   - Calories: [new calorie count]\n"
        "   - Protein: [new protein amount]\n"
        "   - Carbohydrates: [new carbohydrate amount]\n"
        "   - Fiber: [new fiber amount]\n"
        "   - Fat: [new fat amount]\n"
        "   </modified_recipe>\n\n"
        "5. Explanation of Adjustments:\n"
        "   - Provide an explanation of the changes made and any trade-offs or assumptions in 2-3 sentences.\n"
        )

        # Output format
        output_format = (
        "Provide the optimized ingredient quantities based on the above steps.\n"
        "Return the optimized ingredient quantities as a JSON-like string. Strictly follow the format below even if you can't find a solution.\n"
        "If you cannot find a solution, mark `review_needed` as True, explain why it failed specifically in 3-4 sentences in explanation field, and return the original recipe in following format.\n"
        "If you can find a solution (even if not perfect but close), mark `review_needed` as \"\", and provide an explanation of the adjustments.\n"
        "Use only True or \"\" for `review_needed`. Do not change the case.\n"
        "Strictly follow the format including parameter names below and only return the json part, and dont return any other information\n\n"
        "output json:\n"
        '{ "modified_recipe": {"ingredients": ['
        '{"Ingredient Name": "Brussels Sprouts","Component": "veggies", "Grams": 260},'
        '{"Ingredient Name": "Dried Cranberries", "Component": "garnish", "Grams": 2},'
        '{"Ingredient Name": "Pecans", "Component": "garnish", "Grams": 4},'
        '{"Ingredient Name": "Arugula","Component": "veggies", "Grams": 40},'
        '{"Ingredient Name": "Apple Cider Vinegar Dressing","Component": "sauce", "Grams": 20},'
        '{"Ingredient Name": "Carrot", "Component": "veggies", "Grams": 130},'
        '{"Ingredient Name": "Sweet Potato","Component": "starch", "Grams": 200},'
        '{"Ingredient Name": "Chicken Breast","Component": "meat", "Grams": 150}'
        ']},'
        '"updated_nutrition_info": {'
        '"Calories": 628,'
        '"Protein": 48.5,'
        '"Carbohydrates": 64.5,'
        '"Fiber": 15.52,'
        '"Fat": 21.5'
        '},'
        '"explanation": "Increased protein by adding more chicken breast. Adjusted vegetable quantities to meet fiber and carbs goals.",'
        '"review_needed": ""'
        '}'
        )

        # Combine all parts
        return (
        base_description
        + inputs_constraints
        + optional_fields
        + explanation_of_inputs
        + step_by_step
        + output_format
        )

    def call_openai_api(self):
        # Generate the prompt using the generate_prompt function
        prompt = self.generate_prompt()
        
        # Get API key from Streamlit secrets
        api_key = st.secrets["OPENAI_API_KEY"]
        client = OpenAI(api_key=api_key)
        
        # Define the OpenAI API call with the generated prompt
        response = client.chat.completions.create(
            model="gpt-4",  # Fixed typo in model name from "gpt-4o" to "gpt-4"
            messages=[
                {"role": "system", "content": "You are an AI assistant."},
                {"role": "user", "content": prompt}
            ]
        )
        
        ai_response = response.choices[0].message.content
        return ai_response  


    def solve(self):
        prompt = self.generate_prompt()
        response = self.call_openai_api()
        #self.apply_optimization_results(response)
        return response

    def apply_optimization_results(self, response):
        # Assuming the response is a JSON-like string with ingredient quantities
        optimized_ingredients = eval(response)
        for ingredient in self.dish:
            component_name = ingredient['Component (from Ingredient)'][0]
            if component_name in optimized_ingredients:
                scale = optimized_ingredients[component_name]
                for nutrient in self.nutrients:
                    if nutrient in ingredient:
                        ingredient[nutrient] *= scale
                ingredient['Multiplier'] = scale
        for ingredient in self.dish:
            print(f"{ingredient['Component (from Ingredient)'][0]}: {ingredient['Multiplier']}")


if __name__ == "__main__":
    grouped_ingredients="defaultdict(<function MealRecommendation.optimize.<locals>.<lambda> at 0x138357910>, {'veggies': defaultdict(<class 'float'>, {'Kcal': 56.800000000000004, 'Carbohydrate, total (g)': 24.080000000000002, 'Protein (g)': 5.5703, 'Fat, Total (g)': 1.3703, 'Dietary Fiber (g)': 8.05, 'Grams': 300.0}), 'garnish': defaultdict(<class 'float'>, {'Kcal': 19.98, 'Carbohydrate, total (g)': 1.934, 'Protein (g)': 0.18680000000000002, 'Fat, Total (g)': 1.4618, 'Dietary Fiber (g)': 0.298, 'Grams': 4.0}), 'sauce': defaultdict(<class 'float'>, {'Kcal': 96.0, 'Carbohydrate, total (g)': 1.0, 'Protein (g)': 0.0, 'Fat, Total (g)': 10.0, 'Dietary Fiber (g)': 0.0, 'Grams': 20.0}), 'starch': defaultdict(<class 'float'>, {'Kcal': 76.0, 'Carbohydrate, total (g)': 17.7, 'Protein (g)': 1.37, 'Fat, Total (g)': 0.14, 'Dietary Fiber (g)': 2.5, 'Grams': 100.0})})"
    customer_requirements = "{'identifier': 'Richard  Hamilton | Lunch | rejham2@gmail.com', 'First_Name': 'Richard ', 'Last_Name': 'Hamilton', 'goal_calories': 630, 'goal_carbs(g)': 63, 'goal_fiber(g)': 9, 'goal_fat(g)': 21, 'goal_protein(g)': 47, 'Portion Algo Constraints': ['recpZqUPjx32nuwjZ'], 'Meal': 'Lunch', '# of snacks per day': 1, 'Kcal': 630, 'Carbohydrate, total (g)': 63, 'Protein (g)': 47, 'Fat, Total (g)': 21, 'Dietary Fiber (g)': 9}"
    dish="[{'Airtable Dish Name': 'Harvest Bowl', 'Component (from Ingredient)': ['veggies'], 'Ingredient ID': '11745 Brussel Sprouts', 'NDB': '11745', 'Ingredient Name': 'Brussel Sprouts', 'Grams': 130, 'Kcal': 46.800000000000004, 'Protein (g)': 3.315, 'Fat, Total (g)': 0.65, 'Dietary Fiber (g)': 3.3800000000000003, 'Carbohydrate, total (g)': 9.23, 'Recommendation ID': 1, 'Dish ID': 632, 'Client_Id': ['Richard  Hamilton | Lunch | rejham2@gmail.com']}, {'Airtable Dish Name': 'Harvest Bowl', 'Component (from Ingredient)': ['garnish'], 'Ingredient ID': '9079 Dried Cranberries', 'NDB': '9079', 'Ingredient Name': 'Dried Cranberries', 'Grams': 2, 'Kcal': 6.16, 'Protein (g)': 0.0034000000000000002, 'Fat, Total (g)': 0.021800000000000003, 'Dietary Fiber (g)': 0.106, 'Carbohydrate, total (g)': 1.656, 'Recommendation ID': 1, 'Dish ID': 632, 'Client_Id': ['Richard  Hamilton | Lunch | rejham2@gmail.com']}, {'Airtable Dish Name': 'Harvest Bowl', 'Component (from Ingredient)': ['garnish'], 'Ingredient ID': '12142 Pecans', 'NDB': '12142', 'Ingredient Name': 'Pecans', 'Grams': 2, 'Kcal': 13.82, 'Protein (g)': 0.1834, 'Fat, Total (g)': 1.44, 'Dietary Fiber (g)': 0.192, 'Carbohydrate, total (g)': 0.278, 'Recommendation ID': 1, 'Dish ID': 632, 'Client_Id': ['Richard  Hamilton | Lunch | rejham2@gmail.com']}, {'Airtable Dish Name': 'Harvest Bowl', 'Component (from Ingredient)': ['veggies'], 'Ingredient ID': '11959 Arugula', 'NDB': '11959', 'Ingredient Name': 'Arugula', 'Grams': 40, 'Kcal': 10.0, 'Protein (g)': 1.032, 'Fat, Total (g)': 0.264, 'Dietary Fiber (g)': 0.6400000000000001, 'Carbohydrate, total (g)': 1.46, 'Recommendation ID': 1, 'Dish ID': 632, 'Client_Id': ['Richard  Hamilton | Lunch | rejham2@gmail.com']}, {'Airtable Dish Name': 'Harvest Bowl', 'Component (from Ingredient)': ['sauce'], 'Ingredient ID': 'SF Apple Cider Vinegar Dressing', 'NDB': 'SF', 'Ingredient Name': 'Apple Cider Vinegar Dressing', 'Grams': 20, 'Kcal': 96.0, 'Protein (g)': 0.0, 'Fat, Total (g)': 10.0, 'Dietary Fiber (g)': 0.0, 'Carbohydrate, total (g)': 1.0, 'Recommendation ID': 1, 'Dish ID': 632, 'Client_Id': ['Richard  Hamilton | Lunch | rejham2@gmail.com']}, {'Airtable Dish Name': 'Harvest Bowl', 'Component (from Ingredient)': ['veggies'], 'Ingredient ID': '11124 Carrot', 'NDB': '11124', 'Ingredient Name': 'Carrot', 'Grams': 130, 'Kcal': 0.0, 'Protein (g)': 1.2233, 'Fat, Total (g)': 0.4563, 'Dietary Fiber (g)': 4.03, 'Carbohydrate, total (g)': 13.39, 'Recommendation ID': 1, 'Dish ID': 632, 'Client_Id': ['Richard  Hamilton | Lunch | rejham2@gmail.com']}, {'Airtable Dish Name': 'Harvest Bowl', 'Component (from Ingredient)': ['starch'], 'Ingredient ID': '11510 Sweet Potato', 'NDB': '11510', 'Ingredient Name': 'Sweet Potato', 'Grams': 100, 'Kcal': 76.0, 'Protein (g)': 1.37, 'Fat, Total (g)': 0.14, 'Dietary Fiber (g)': 2.5, 'Carbohydrate, total (g)': 17.7, 'Recommendation ID': 1, 'Dish ID': 632, 'Client_Id': ['Richard  Hamilton | Lunch | rejham2@gmail.com']}]"
    nutrients="['Kcal', 'Carbohydrate, total (g)', 'Protein (g)', 'Fat, Total (g)', 'Dietary Fiber (g)', 'Grams']"
    nutrient_constraints="{'Protein (g)': {'lb': 0.8, 'ub': 1.2}, 'Kcal': {'lb': 0.8, 'ub': 1.2}, 'Fat, Total (g)': {'lb': 0.8, 'ub': 1.2}, 'Dietary Fiber (g)': {'lb': 0.8, 'ub': 1.2}, 'Carbohydrate, total (g)': {'lb': 0.7, 'ub': 1.2}}"
    garnish_grams=4
    sauce_grams=20
    veggie_ge_starch=True
    min_meat_per_100_cal=10
    max_meal_grams_per_100_cal=200
    optimizer = LLMDishOptimizer(grouped_ingredients, customer_requirements, nutrients, nutrient_constraints, garnish_grams, sauce_grams, veggie_ge_starch, min_meat_per_100_cal, max_meal_grams_per_100_cal,dish)
    result = optimizer.solve()
    print(result)
    
    