import numpy as np
import random
import math

SPECIAL_YOGURT_PROTEIN_ITEM_KEYWORDS = ["overnight oats", "yogurt", "yoghurt","parfait"]
MAX_SPECIAL_YOGURT_PROTEIN_GRAM = 400
MAX_SPECIAL_YOGURT_PROTEIN_ITEM_TOTAL_MEAL = 500
MAX_SPECIAL_YOGURT_VEGGIES_GRAM = 20
SPECIAL_FRUIT_SNACK_DISH_KEYWORDS = ['seasonal fruit salad']
MAX_SPECIAL_FRUIT_SNACK_DISH_VEGGIES_GRAM = 220
MAX_VEGGIES_GRAM = 300
MAX_STARCH_GRAM = 280
MIN_STARCH_GRAM = 50
MAX_PROTEIN_PER_TYPE = {"meat": 200, "fish": 220, "tofu": 350, "vegan": 200}

class NewDishOptimizer:
    def __init__(self, grouped_ingredients, customer_requirements, nutrients, nutrient_constraints,
                 garnish_grams=None, double_sauce=False, veggie_ge_starch=True, 
                 min_meat_per_100_cal=None, max_meal_grams_per_100_cal=None, dish=None):
        """ print(f"Initialized Dish Optimizer with {len(grouped_ingredients)} ingredients/n")
        print(f"grouped_ingredients: {grouped_ingredients}\n")
        print(f"Customer Requirements: {customer_requirements}\n")
        print(f"Nutrient: {nutrients}\n")
        print(f"Nutrient Constraints: {nutrient_constraints}\n")
        print(f"Garnish Grams: {garnish_grams}\n")
        print(f"Sauce Grams: {sauce_grams}\n")
        print(f"Veggie >= Starch: {veggie_ge_starch}\n")
        print(f"Min Meat per 100 cal: {min_meat_per_100_cal}\n")
        print(f"Max Meal Grams per 100 cal: {max_meal_grams_per_100_cal}\n")
        print(f"Dish: {dish}\n")
 """
        self.grouped_ingredients = grouped_ingredients
        self.customer_requirements = self._normalize_requirements(customer_requirements)
        self.nutrients = nutrients
        self.nutrient_constraints = nutrient_constraints
        self.garnish_grams = garnish_grams
        #self.sauce_grams = sauce_grams
        self.double_sauce = double_sauce
        self.veggie_ge_starch = veggie_ge_starch
        self.min_meat_per_100_cal = min_meat_per_100_cal
        self.max_meal_grams_per_100_cal = max_meal_grams_per_100_cal
        self.dish = self._initialize_recipe(dish) if dish else None
        self.scaler_penalty_weight = 0.1  # Penalty weight for scaler deviations
        self.is_special_yogurt_protein = False  # Flag to track if we're processing a special protein item
        self.is_special_fruit_snack = False  # Flag to track if we're processing a special fruit snack
        
        # Simplified base weights - only keep essential ones
        # Modified base weights - increased protein priority and fat penalty
        self.nutrient_weights = {
            'kcal': 5,        # Balanced priority for calories
            'protein(g)': 5,  # Balanced protein priority
            'carbohydrate(g)': 3,  # Balanced carb priority
            'dietaryFiber(g)': 4,
            'fat(g)': 1       # Balanced fat priority
        }

        
    def get_bound_based_ratio(self, nutrient, current_value):
        """
        Calculate a ratio of how far the current value is from the nutrient's bounds.
        Returns:
        - 1.0 if within bounds
        - > 1.0 if above upper bound (proportional to exceedance)
        - < 1.0 if below lower bound (inversely proportional to deficit)
        """
        # Retrieve target value for the nutrient
        target = self.customer_requirements.get(nutrient, 0)
        if target <= 0:
            return 1.0

        # Get constraints
        normalized_nutrient = self.normalize_nutrient_name(nutrient)
        constraint = next(
            (bounds for k, bounds in self.nutrient_constraints.items() 
            if self.normalize_nutrient_name(k) == normalized_nutrient),
            None
        )

        if not constraint:
            return 1.0

        # Calculate bounds
        upper_bound = constraint.get('ub')
        if upper_bound is None:
            upper_bound = float('inf')
        else:
            upper_bound *= target

        lower_bound = constraint.get('lb')
        if lower_bound is None:
            lower_bound = 0.00000000000001
        else:
            lower_bound *= target


        """ # Ratio calculation logic:
            # - If the current value exceeds the upper bound, ratio > 1.0 (proportional to excess)
            # - If the current value is below the target (or lower bound for non-protein nutrients), ratio < 1.0 (inversely proportional to deficit)
            # - If within target range, return 1.0 (no adjustment needed)
        if current_value > upper_bound:
            # Over target: ratio > 1.0, proportional to exceedance
            return current_value / upper_bound
        elif current_value < (target if normalized_nutrient == 'protein' else lower_bound):
            # Under target: ratio < 1.0, inversely proportional to deficit
            return current_value / (target if normalized_nutrient == 'protein' else lower_bound)
        else:
            # Within bounds: return default ratio
            return 1.0 """
        return current_value / target    

    def _get_diff_ratios(self, current_nutrition):
        """Calculate basic ratios based on bounds"""
        ratios = {}
        for nutrient in self.nutrients:
            if nutrient in self.customer_requirements:
                ratios[nutrient] = self.get_bound_based_ratio(nutrient, current_nutrition[nutrient])
        return ratios

    
    def normalize_nutrient_name(self, nutrient):
        """Normalize nutrient names to handle different formats."""
        return nutrient.lower().replace('(g)', '').replace('total', '').replace(' ', '').replace(',', '')

    def _initialize_recipe(self, dish):
        """Initialize recipe with default scalers."""
        if not dish or 'ingredients' not in dish:
            return dish
        return {
            **dish,
            'ingredients': [{**ingredient, 'scaler': 1.0} for ingredient in dish['ingredients']]
        }

    def _normalize_requirements(self, requirements):
        """Normalize customer requirements to standard format."""
        key_mapping = {
            'goal_calories': 'kcal',
            'Kcal': 'kcal',
            'goal_protein(g)': 'protein(g)',
            'Protein (g)': 'protein(g)',
            'goal_fat(g)': 'fat(g)',
            'Fat, Total (g)': 'fat(g)',
            'goal_fiber(g)': 'dietaryFiber(g)',
            'Dietary Fiber (g)': 'dietaryFiber(g)',
            'goal_carbs(g)': 'carbohydrate(g)',
            'Carbohydrate, total (g)': 'carbohydrate(g)'
        }
        
        return {
            key_mapping[key]: float(value)
            for key, value in requirements.items()
            if key in key_mapping and value is not None
        }

    def get_effective_grams(self, ingredient):
        """Calculate effective grams based on base grams and scaler."""
        return ingredient['baseGrams'] * ingredient.get('scaler', 1.0)

    def calculate_total_nutrition(self, recipe):
        """Calculate total nutrition values for the recipe."""
        totals = {nutrient: 0 for nutrient in self.nutrients}
        
        for ingredient in recipe:
            try:
                effective_grams = self.get_effective_grams(ingredient)
                if effective_grams <= 0:
                    continue
                    
                for nutrient in self.nutrients:
                    nutrient_value = ingredient.get(f'{nutrient}PerBaseGrams', 0)
                    if nutrient_value is not None:
                        totals[nutrient] += float(nutrient_value) * ingredient.get('scaler', 1.0)
                        
            except (ValueError, TypeError) as e:
                print(f"Warning: Error processing {ingredient.get('ingredientName', 'unknown')}: {str(e)}")
                
        return totals

    def calculate_weighted_deviation(self, current, target, recipe=None):
        """Calculate weighted deviation with bound-based comparisons"""
        total_deviation = 0
        
        # Calculate nutrition deviation
        for nutrient, value in current.items():
            if nutrient in target and target[nutrient] > 0:
                weight = self.nutrient_weights.get(nutrient, 1.0)
                
                # Find matching constraint
                normalized_nutrient = self.normalize_nutrient_name(nutrient)
                constraint = next(
                    (bounds for k, bounds in self.nutrient_constraints.items() 
                    if self.normalize_nutrient_name(k) == normalized_nutrient),
                    None
                )
                
                if constraint:
                    # Get bounds
                    lower_bound = constraint.get('lb', 1.0) * target[nutrient]
                    upper_bound = constraint.get('ub', 1.0)
                    if upper_bound is None:
                        upper_bound = float('inf')
                    upper_bound *= target[nutrient]
                    
                    # Calculate deviation based on which bound is violated
                    if value > target[nutrient]:
                        # If over target, compare with upper bound
                        relative_dev = (value - upper_bound) / upper_bound if value > upper_bound else 0
                    else:
                        # If under target, compare with lower bound
                        relative_dev = (value - lower_bound) / lower_bound if value < lower_bound else 0
                else:
                    # If no constraint, use target directly
                    relative_dev = (value - target[nutrient]) / target[nutrient]
                    
                total_deviation += weight * relative_dev ** 2

        # Add penalty for sum of scalers
        if recipe:
            total_scaler = sum(abs(ing.get('scaler', 1.0) - 1.0) for ing in recipe)
            total_deviation += self.scaler_penalty_weight * total_scaler

        return total_deviation


    def check_recipe_constraints(self, recipe):
        """
        Check if recipe meets all additional constraints.
        Returns True if all constraints are met, False otherwise.
        """
        try:
            # Check sauce constraint
            """ if self.sauce_grams is not None:
                sauce_total = sum(self.get_effective_grams(i) for i in recipe if i['component'] == 'sauce')
                if sauce_total != self.sauce_grams:
                    # print(f"Sauce constraint violation: {sauce_total:.1f}g vs required {self.sauce_grams}g")
                    return False """

            # Check veggie >= starch constraint
            if self.veggie_ge_starch:
                veggies_total = sum(self.get_effective_grams(i) for i in recipe if i['component'] == 'veggies')
                starch_total = sum(self.get_effective_grams(i) for i in recipe if i['component'] == 'starch')
                if veggies_total < starch_total or veggies_total > MAX_VEGGIES_GRAM or starch_total > MAX_STARCH_GRAM:
                    # print(f"Veggie-starch constraint violation: veggies={veggies_total:.1f}g, starch={starch_total:.1f}g")
                    return False

            # Check minimum meat per 100 cal constraint
            if self.min_meat_per_100_cal:
                nutrition = self.calculate_total_nutrition(recipe)
                total_kcal = nutrition['kcal']
                if total_kcal > 0:
                    total_meat = sum(self.get_effective_grams(i) for i in recipe if 'protein' in i['component'])
                    meat_per_100_cal = (total_meat / total_kcal) * 100
                    if meat_per_100_cal < self.min_meat_per_100_cal:
                        # print(f"Meat per 100 cal constraint violation: {meat_per_100_cal:.1f} < {self.min_meat_per_100_cal}")
                        return False

            if self.is_special_yogurt_protein:
                total_grams = sum(self.get_effective_grams(i) for i in recipe)
                if total_kcal > 0:
                    if total_grams > MAX_SPECIAL_YOGURT_PROTEIN_ITEM_TOTAL_MEAL:
                        return False
            # Check maximum meal grams per 100 cal constraint
            elif self.max_meal_grams_per_100_cal:
                total_grams = sum(self.get_effective_grams(i) for i in recipe)
                nutrition = self.calculate_total_nutrition(recipe)
                total_kcal = nutrition['kcal']
                if total_kcal > 0:
                    grams_per_100_cal = (total_grams / total_kcal) * 100
                    if grams_per_100_cal > self.max_meal_grams_per_100_cal:
                        # print(f"Grams per 100 cal constraint violation: {grams_per_100_cal:.1f} > {self.max_meal_grams_per_100_cal}")
                        return False
            
            return True  # All constraints met
            
        except Exception as e:
            print(f"Error in checking constraints: {str(e)}")
            return False 

    def is_within_nutrition_range(self, recipe, nutrition_totals):
        """Check if recipe meets all nutritional constraints."""
        for nutrient, value in nutrition_totals.items():
            value = round(value, 1)
            normalized_nutrient = self.normalize_nutrient_name(nutrient)
            
            matching_constraints = [
                (k, v) for k, v in self.nutrient_constraints.items() 
                if self.normalize_nutrient_name(k) == normalized_nutrient
            ]
            
            for constraint_key, bounds in matching_constraints:
                target = self.customer_requirements[nutrient]
                
                if bounds.get('lb') and value < round(bounds['lb'] * target, 1):
                    # print(f"Constraint violation - {nutrient} too low: {value:.1f} < {bounds['lb'] * target:.1f}")
                    return False
                    
                if bounds.get('ub') and value > round(bounds['ub'] * target, 1):
                    # print(f"Constraint violation - {nutrient} too high: {value:.1f} > {bounds['ub'] * target:.1f}")
                    return False
        
        if self.max_meal_grams_per_100_cal and nutrition_totals['kcal'] > 0:
            total_grams = sum(self.get_effective_grams(i) for i in recipe)
            grams_per_100_cal = (total_grams / nutrition_totals['kcal']) * 100
            if grams_per_100_cal > self.max_meal_grams_per_100_cal:
                # print(f"Constraint violation - grams per 100 cal too high: {grams_per_100_cal:.1f} > {self.max_meal_grams_per_100_cal}")
                return False
        
        # check total grams of starch
        if sum(self.get_effective_grams(i) for i in recipe if i['component'] == 'starch') < MAX_STARCH_GRAM:
            return False
        
        return True

    def _calculate_ingredient_contributions(self, recipe):
        """Calculate each ingredient's contribution to each nutrient per gram."""
        contributions = []
        for idx, ingredient in enumerate(recipe):
            nutrient_density = {}
            for nutrient in self.nutrients:
                per_gram = ingredient.get(f'{nutrient}PerBaseGrams', 0) / ingredient['baseGrams'] if ingredient['baseGrams'] > 0 else 0
                if per_gram > 0:
                    nutrient_density[nutrient] = per_gram
            contributions.append((idx, nutrient_density))
        return contributions

    def format_result(self, recipe, nutrition_totals):
        """Format the optimization results into the required JSON structure."""
        # Format modified recipe
        modified_recipe = {
            "ingredients": [
                {
                    "ingredientName": f"{ing['ingredientName']} ({ing.get('scaler')} x sauce)" if ing['component'] == 'sauce' and ing.get('scaler', 1.0) != 1.0 else ing['ingredientName'],
                    "Component": ing['component'],
                    "Grams": round(ing['baseGrams'] * ing.get('scaler', 1.0), 2),
                    "ingredientId": ing['ingredientId']
                }
                for ing in recipe
            ]
        }
        
        # Calculate percentages
        percentages = {}
        for nutrient in ['kcal', 'protein(g)', 'fat(g)', 'dietaryFiber(g)', 'carbohydrate(g)']:
            if self.customer_requirements[nutrient] != 0:
                percentages[nutrient] = round(nutrition_totals[nutrient] / self.customer_requirements[nutrient] * 100, 1)
            else:
                percentages[nutrient] = 0

        # Format explanation string
        explanation = (
            f"\nAchieved / Target Nutrition:\n"
            f"kcal: {nutrition_totals['kcal']:.1f} / {self.customer_requirements['kcal']:.1f} "
            f"({percentages['kcal']:.1f}% of target)\n"
            f"protein(g): {nutrition_totals['protein(g)']:.1f} / {self.customer_requirements['protein(g)']:.1f} "
            f"({percentages['protein(g)']:.1f}% of target)\n"
            f"fat(g): {nutrition_totals['fat(g)']:.1f} / {self.customer_requirements['fat(g)']:.1f} "
            f"({percentages['fat(g)']:.1f}% of target)\n"
            f"dietaryFiber(g): {nutrition_totals['dietaryFiber(g)']:.1f} / {self.customer_requirements['dietaryFiber(g)']:.1f} "
            f"({percentages['dietaryFiber(g)']:.1f}% of target)\n"
            f"carbohydrate(g): {nutrition_totals['carbohydrate(g)']:.1f} / {self.customer_requirements['carbohydrate(g)']:.1f} "
            f"({percentages['carbohydrate(g)']:.1f}% of target)"
        )
        
        # Format updated nutrition info
        updated_nutrition_info = {
            "Calories": round(nutrition_totals['kcal'], 1),
            "Protein": round(nutrition_totals['protein(g)'], 1),
            "Carbohydrates": round(nutrition_totals['carbohydrate(g)'], 1),
            "Fiber": round(nutrition_totals['dietaryFiber(g)'], 1),
            "Fat": round(nutrition_totals['fat(g)'], 1),
            # Add percentage values
            "Calories %": percentages['kcal'],
            "Protein %": percentages['protein(g)'],
            "Fat %": percentages['fat(g)'],
            "Fiber %": percentages['dietaryFiber(g)'],
            "Carbs %": percentages['carbohydrate(g)'],
            "Sodium (mg)": round(nutrition_totals.get('Sodium (mg)', 0), 1),
            "Calcium (mg)": round(nutrition_totals.get('Calcium (mg)', 0), 1),
            "Phosphorus, P (mg)": round(nutrition_totals.get('Phosphorus, P (mg)', 0), 1),
            "Fatty acids, total saturated (g)": round(nutrition_totals.get('Fatty acids, total saturated (g)', 0), 1)
        }
        
        # Check constraints and prepare results
        meets_constraints = self.is_within_nutrition_range(recipe, nutrition_totals)
        constraint_violations = []
        
        # Check each nutrient constraint
        for nutrient, value in nutrition_totals.items():
            value = round(value, 1)
            normalized_nutrient = self.normalize_nutrient_name(nutrient)
            constraint_key = next((k for k in self.nutrient_constraints.keys() 
                                if self.normalize_nutrient_name(k) == normalized_nutrient), None)
            
            if constraint_key:
                bounds = self.nutrient_constraints[constraint_key]
                target = self.customer_requirements[nutrient]
                # print(f"Checking {nutrient} with value {value} and target {target} and bounds {bounds}")

                if bounds.get('lb') and value < round(bounds['lb'] * target, 1):
                    constraint_violations.append(
                        f"{nutrient} too low: {value:.1f} < {bounds['lb'] * target:.1f}"
                    )
                if bounds.get('ub') and value > round(bounds['ub'] * target, 1):
                    constraint_violations.append(
                        f"{nutrient} too high: {value:.1f} > {bounds['ub'] * target:.1f}"
                    )
        
        results = {
            "review_needed": not meets_constraints,
            "notes": (
                "All constraints met" if meets_constraints else 
                "; ".join(constraint_violations)
            )
        }
        
        return {
            "modified_recipe": {
                "ingredients": modified_recipe["ingredients"]
            },
            "explanation": explanation,
            "updated_nutrition_info": updated_nutrition_info,
            "results": results
        }

    def get_component_grams(self, recipe, component):
        """Calculate total grams for a specific component."""
        return sum(self.get_effective_grams(ing) for ing in recipe if ing['component'] == component)

    def adjust_component_within_limit(self, recipe, component, max_grams):
        """Adjust component scalers to stay within limit."""
        total_grams = self.get_component_grams(recipe, component)
        if total_grams > max_grams:
            reduction_factor = max_grams / total_grams
            for idx, ing in enumerate(recipe):
                if ing['component'] == component:
                    recipe[idx]['scaler'] *= reduction_factor
        return recipe
    
    def adjust_component_above_minimum(self, recipe, component, min_grams):
        """Adjust component scalers to meet minimum gram requirement."""
        total_grams = self.get_component_grams(recipe, component)
        if total_grams < min_grams:
            # Calculate how much we need to increase
            increase_factor = min_grams / total_grams
            # Apply increase to all ingredients of this component
            for idx, ing in enumerate(recipe):
                if ing['component'] == component:
                    recipe[idx]['scaler'] *= increase_factor
        return recipe
    
    def _is_valid_adjustment(self, recipe, idx, new_scaler):
        """New helper function to check if adjustment maintains constraints"""
        component = recipe[idx]['component']
        
        if component == 'veggies':
            potential_grams = sum(
                self.get_effective_grams(x) if i != idx else x['baseGrams'] * new_scaler
                for i, x in enumerate(recipe) if x['component'] == 'veggies'
            )
            if self.is_special_yogurt_protein:
                if potential_grams > MAX_SPECIAL_YOGURT_VEGGIES_GRAM:
                    return False
            elif self.is_special_fruit_snack:
                if potential_grams > MAX_SPECIAL_FRUIT_SNACK_DISH_VEGGIES_GRAM:
                    return False
            elif potential_grams > MAX_VEGGIES_GRAM:
                return False
                
        elif component == 'starch':
            potential_grams = sum(
                self.get_effective_grams(x) if i != idx else x['baseGrams'] * new_scaler
                for i, x in enumerate(recipe) if x['component'] == 'starch'
            )
            if potential_grams > MAX_STARCH_GRAM:
                return False
        
        elif component == 'protein' and self.is_special_yogurt_protein:
            potential_grams = recipe[idx]['baseGrams'] * new_scaler
            if potential_grams > MAX_SPECIAL_YOGURT_PROTEIN_GRAM:
                return False
        
        return True

    def _get_ingredient_adjustment(self, component, ratios, contributions):
        """
        Calculate ingredient adjustments following nutritional priorities with clear hierarchy:
        1. Protein optimization (highest priority)
            - Balance protein needs while considering calorie and carb impact
        2. Fiber requirements through veggie adjustments
            - Manage fiber levels while considering calorie and carb balance
        3. Overall caloric balance through starch adjustments
            - Fine-tune total calories while maintaining macro ratios
        
        Parameters:
        - component: str - Type of ingredient ('protein', 'veggies', or 'starch')
        - ratios: dict - Current nutrient ratios relative to target range (current/target)
        - contributions: dict - Nutrient contributions from each component
        
        Returns:
        - float: Adjustment factor with slight randomization (negative reduces, positive increases)
        """    
        # Extract ratios with defaults
        protein_ratio = ratios.get('protein(g)', 1.0)
        kcal_ratio = ratios.get('kcal', 1.0)
        carb_ratio = ratios.get('carbohydrate(g)', 1.0)
        fiber_ratio = ratios.get('dietaryFiber(g)', 1.0)
        
        # Get contributions for comparing protein vs carb balance
        protein_contrib = contributions.get('protein(g)', 0)
        carb_contrib = contributions.get('carbohydrate(g)', 0)
        
        # Initialize adjustment (negative reduces portion, positive increases)
        adj = 0

        if component == 'protein':
            if protein_ratio > 1.0:
                base_adj = ((1.0 / protein_ratio) - 1.0)
                if kcal_ratio < 1.0:
                    adj = min(base_adj, ((1.0 / kcal_ratio) - 1.0)/5)
                elif carb_ratio > 1 and carb_contrib > protein_contrib:
                    adj = base_adj * 1.1
                else:
                    adj = base_adj * 1.2
            elif protein_ratio < 1.0:
                base_adj = ((1.0 / protein_ratio) - 1.0)
                if kcal_ratio < 1.0:
                    adj = max(base_adj * 0.7, ((1.0 / kcal_ratio) - 1.0)/4)
                elif carb_ratio > 1:
                    adj = base_adj * 0.6
                else:
                    adj = base_adj
            elif protein_ratio > 1 and protein_ratio < 1.03 and kcal_ratio < 1.0:
                adj = ((1.0 / kcal_ratio) - 1.0) / 4
                    
        elif component == 'veggies':
            if fiber_ratio < 1:
                base_adj = (1.0 / fiber_ratio) - 1.0
                if carb_ratio > 1.2:
                    adj = base_adj * 0.7  
                else:
                    adj = base_adj  
            elif fiber_ratio > 1 and kcal_ratio > 1:
                base_adj = ((1.0 / kcal_ratio) - 1.0) / 4 
                if carb_ratio > 1.0:
                    adj = base_adj * 1.2  
                else:
                    adj = base_adj
            elif carb_ratio < 1.0:
                adj = ((1.0 / carb_ratio) - 1.0) / 4
                
        elif component == 'starch':
            # Modified starch adjustments for more aggressive increases
            if carb_ratio > 1.0:
                # Less aggressive reduction for high carbs
                base_reduction = ((1.0 / carb_ratio) - 1.0)
                if kcal_ratio < 1.0:
                    adj = base_reduction * 0.6  # Less reduction (was 0.8)
                else:
                    adj = base_reduction * 1.1  # Less reduction (was 1.3)
            elif carb_ratio < 1.0:
                # More aggressive increase for low carbs
                base_adj = (1.0 / carb_ratio) - 1.0
                if kcal_ratio > 1.0:
                    adj = base_adj * 0.3  
                elif protein_ratio < 1.0:
                    adj = base_adj * 0.2  
                else:
                    adj = base_adj * 0.6 
            elif kcal_ratio < 1.0:
                # More aggressive increase for low calories
                adj = ((1.0 / kcal_ratio) - 1.0) / 4
            elif kcal_ratio > 1.0:
                # Less aggressive reduction for high calories
                adj = ((1.0 / kcal_ratio) - 1.0) / 4

            if protein_contrib / carb_contrib > 0.5:
                if adj > 0:
                    adj *= 0.3
                else:
                    adj *= 2

        # Add small randomization to prevent getting stuck in local optima
        return adj
    
    def _adjust_ingredients_sequentially(self, recipe, initial_nutrition):
        """Adjust ingredients one at a time following manual optimization steps"""
        adjusted_recipe = [{**ing} for ing in recipe]
        current_nutrition = initial_nutrition.copy()
        
        # Define adjustment sequence based on manual steps
        component_sequence = [
            # Step 1: Protein first
            [(i, ing) for i, ing in enumerate(recipe) if ing['component'] == 'protein'],
            # Step 2: Veggies for fiber
            [(i, ing) for i, ing in enumerate(recipe) if ing['component'] == 'veggies'],
            # Step 3-6: Starch for calorie management
            [(i, ing) for i, ing in enumerate(recipe) if ing['component'] == 'starch']
        ]
        
        # Process each component group in sequence
        for component_group in component_sequence:
            for idx, ingredient in component_group:
                # Calculate current nutrient gaps
                nutrient_gaps = {
                    nutrient: (self.customer_requirements[nutrient] - current_nutrition[nutrient]) / self.customer_requirements[nutrient] if self.customer_requirements[nutrient] != 0 else 0
                    for nutrient in self.nutrients if nutrient in self.customer_requirements
                }
                
                # Calculate adjustment for single ingredient
                adjustment = self._calculate_single_adjustment(adjusted_recipe[idx], nutrient_gaps, current_nutrition)
                
                if adjustment != 0:
                    # Apply adjustment
                    current_scaler = adjusted_recipe[idx].get('scaler', 1.0)
                    new_scaler = max(0.1, current_scaler * (1.0 + adjustment))
                    
                    if self._is_valid_adjustment(adjusted_recipe, idx, new_scaler):
                        # Update scaler and recalculate nutrition
                        adjusted_recipe[idx]['scaler'] = math.ceil(new_scaler * 100)/100
                        current_nutrition = self.calculate_total_nutrition(adjusted_recipe)
        
        return adjusted_recipe

    def _calculate_single_adjustment(self, ingredient, nutrient_gaps, current_nutrition):
        """Calculate adjustment for a single ingredient based on current nutritional state"""
        component = ingredient['component']
        
        if component in ['sauce', 'garnish'] or ingredient['baseGrams'] <= 0:
            return 0
            
        # Calculate nutrient contributions
        contributions = {
            nutrient: ingredient.get(f'{nutrient}PerBaseGrams', 0) / ingredient['baseGrams']
            for nutrient in nutrient_gaps.keys()
            if ingredient.get(f'{nutrient}PerBaseGrams', 0) > 0
        }
        
        if not contributions:
            return 0
            
        nutrient_ratios = self._get_diff_ratios(current_nutrition)
        
        # Use existing adjustment logic but for single ingredient
        base_adjustment = self._get_ingredient_adjustment(component, nutrient_ratios, contributions)
        return base_adjustment

    def _recipes_are_similar(self, recipe1, recipe2, threshold=0.01):
        """Check if two recipes are similar enough to consider converged"""
        for ing1, ing2 in zip(recipe1, recipe2):
            if abs(ing1.get('scaler', 1.0) - ing2.get('scaler', 1.0)) > threshold:
                return False
        return True
    
    def _final_adjustment(self, formatted_result, final_recipe, final_nutrition):
        notes = formatted_result['results']['notes']
        
        # Helper functions for common operations with more robust parsing
        def extract_values(note_type):
            """Extract current and target values from a specific note type."""
            if note_type in notes:
                # Extract the specific note segment by finding the pattern
                pattern = f"{note_type}: "
                start_pos = notes.find(pattern)
                if start_pos == -1:
                    return 0
                
                # Find the semicolon that ends this note or end of string
                end_pos = notes.find(';', start_pos)
                if end_pos == -1:
                    end_pos = len(notes)
                
                note_segment = notes[start_pos + len(pattern):end_pos]
                
                # Split to get current and target values based on comparison operator
                if ' < ' in note_segment:  # Too low case
                    parts = note_segment.split(' < ')
                    if len(parts) != 2:
                        return 0
                    current = float(parts[0].strip())
                    target = float(parts[1].strip())
                    return target - current  # Positive means need to add more
                elif ' > ' in note_segment:  # Too high case
                    parts = note_segment.split(' > ')
                    if len(parts) != 2:
                        return 0
                    current = float(parts[0].strip())
                    target = float(parts[1].strip())
                    return target - current  # Negative means need to reduce
                
            return 0
        
        def get_carbs_needed():
            low_value = extract_values("carbohydrate(g) too low")
            high_value = extract_values("carbohydrate(g) too high")
            return low_value or high_value  # Return whichever is non-zero
            
        def get_protein_needed():
            low_value = extract_values("protein(g) too low")
            high_value = extract_values("protein(g) too high")
            return low_value or high_value  # Return whichever is non-zero
            
        def get_kcal_needed():
            low_value = extract_values("kcal too low")
            high_value = extract_values("kcal too high")
            return low_value or high_value  # Return whichever is non-zero
            
        def add_veggies_for_carbs(carbs_needed):
            # Try to add veggies to increase carbs
            total_veggie_grams = sum(self.get_effective_grams(item) for item in final_recipe if item["component"] == "veggies")
            remaining_veggie_capacity = MAX_VEGGIES_GRAM - total_veggie_grams
            
            if remaining_veggie_capacity > 0:
                # Find veggie with highest carb content
                veggies = [item for item in final_recipe if item["component"] == "veggies"]
                if veggies:
                    highest_carb_veggie = max(
                        veggies, 
                        key=lambda x: x["carbohydrate(g)PerBaseGrams"]
                    )
                    
                    # Calculate how many carbs we can add through veggies
                    carbs_per_gram = highest_carb_veggie["carbohydrate(g)PerBaseGrams"] / highest_carb_veggie["baseGrams"]
                    grams_to_add = min(remaining_veggie_capacity, carbs_needed / carbs_per_gram)
                    highest_carb_veggie["scaler"] += grams_to_add / highest_carb_veggie["baseGrams"]
                    return True
            return False
            
        def add_starch_for_carbs(carbs_needed):
            # Try to add starch to increase carbs
            starch_items = [item for item in final_recipe if item["component"] == "starch"]
            if starch_items:
                starch_item = starch_items[0]
                current_starch_grams = sum(self.get_effective_grams(item) for item in final_recipe 
                                        if item["component"] == "starch")
                remaining_starch_capacity = MAX_STARCH_GRAM - current_starch_grams
                
                if remaining_starch_capacity > 0:
                    carbs_per_gram = starch_item["carbohydrate(g)PerBaseGrams"] / starch_item["baseGrams"]
                    grams_to_add = min(remaining_starch_capacity, carbs_needed / carbs_per_gram)
                    starch_item["scaler"] += grams_to_add / starch_item["baseGrams"]
                    return True
            return False
            
        def add_protein(protein_needed, for_calories=False):
            # Try to add protein to increase protein or calories
            protein_items = [item for item in final_recipe if item["component"] == "protein"]
            if protein_items:
                protein_item = protein_items[0]
                
                # Determine how much to add based on protein or calories needed
                if for_calories:
                    kcal_needed = get_kcal_needed()
                    kcal_per_gram = protein_item["kcalPerBaseGrams"] / protein_item["baseGrams"]
                    grams_to_add = kcal_needed / kcal_per_gram
                else:
                    protein_per_gram = protein_item["protein(g)PerBaseGrams"] / protein_item["baseGrams"]
                    grams_to_add = protein_needed / protein_per_gram
                
                # Check if it's a special protein item
                if (any(item.lower() in protein_item["ingredientName"].lower() 
                    for item in SPECIAL_YOGURT_PROTEIN_ITEM_KEYWORDS)):
                        current_grams = self.get_effective_grams(protein_item)
                        max_additional_grams = max(0, MAX_SPECIAL_YOGURT_PROTEIN_GRAM - current_grams)
                        grams_to_add = min(grams_to_add, max_additional_grams)
                else:
                    # Regular protein item
                    remaining_protein_capacity = self.protein_max - sum(self.get_effective_grams(item) for item in final_recipe if item["component"] == "protein")
                    if remaining_protein_capacity > 0:
                        grams_to_add = min(grams_to_add, remaining_protein_capacity)
                    else:
                        return False
                        
                protein_item["scaler"] += grams_to_add / protein_item["baseGrams"]
                return True
            return False
            
        def double_sauce():
            # Try doubling sauce for more calories
            sauce_items = [item for item in final_recipe if item["component"] == "sauce"]
            if sauce_items:
                adjusted = False
                for sauce_item in sauce_items:
                    current_scaler = int(sauce_item["scaler"])  # Ensure current scaler is integer
                    if current_scaler == 1:  # Only double if currently at 1
                        sauce_item["scaler"] = 2
                        adjusted = True
                return adjusted
            return False
            
        def reduce_protein(excess_protein):
            """Reduce protein when it's too high."""
            protein_items = [item for item in final_recipe if item["component"] == "protein"]
            if protein_items and excess_protein < 0:  # excess_protein will be negative when too high
                protein_item = protein_items[0]
                current_grams = self.get_effective_grams(protein_item)
                
                # Calculate how much to reduce while ensuring we don't go below MIN_PROTEIN_GRAM
                protein_per_gram = protein_item["protein(g)PerBaseGrams"] / protein_item["baseGrams"]
                grams_to_reduce = abs(excess_protein) / protein_per_gram
                
                # Ensure we don't reduce below minimum
                min_protein_grams = MIN_PROTEIN_GRAM if hasattr(self, 'MIN_PROTEIN_GRAM') else current_grams * 0.5  # fallback to 50% reduction if no min defined
                max_reducible = current_grams - min_protein_grams
                grams_to_reduce = min(grams_to_reduce, max_reducible)
                
                if grams_to_reduce > 0:
                    scaler_reduction = grams_to_reduce / protein_item["baseGrams"]
                    protein_item["scaler"] = max(0.1, protein_item["scaler"] - scaler_reduction)  # Ensure we don't go below 0.1 scaler
                    return True
            return False
        
        def reduce_carbs(excess_carbs):
            """Reduce carbs when they're too high, first starch then veggies."""
            if excess_carbs >= 0:  # Not excess
                return False
                
            # First try to reduce starch
            starch_items = [item for item in final_recipe if item["component"] == "starch"]
            reduced = False
            
            if starch_items:
                starch_item = starch_items[0]
                current_grams = self.get_effective_grams(starch_item)
                carbs_per_gram = starch_item["carbohydrate(g)PerBaseGrams"] / starch_item["baseGrams"]
                
                # Calculate how much we need to reduce
                grams_to_reduce = abs(excess_carbs) / carbs_per_gram
                
                # Ensure we don't reduce below minimum starch amount
                min_starch_grams = MIN_STARCH_GRAM if hasattr(self, 'MIN_STARCH_GRAM') else current_grams * 0.5  # fallback to 50%
                max_reducible = current_grams - min_starch_grams
                grams_to_reduce = min(grams_to_reduce, max_reducible)
                
                if grams_to_reduce > 0:
                    scaler_reduction = grams_to_reduce / starch_item["baseGrams"]
                    starch_item["scaler"] = max(0.1, starch_item["scaler"] - scaler_reduction)
                    reduced = True
            
            # Recalculate and see if we still need to reduce carbs
            if reduced:
                current_nutrition = self.calculate_total_nutrition(final_recipe)
                notes = self.format_result(final_recipe, current_nutrition)['results']['notes']
                excess_carbs = get_carbs_needed()
            
            # If still need to reduce carbs, try reducing veggies
            if excess_carbs < 0:
                veggie_items = [item for item in final_recipe if item["component"] == "veggies"]
                if veggie_items:
                    # Find veggie with highest carb content to reduce
                    highest_carb_veggie = max(
                        veggie_items, 
                        key=lambda x: x["carbohydrate(g)PerBaseGrams"]
                    )
                    
                    current_grams = self.get_effective_grams(highest_carb_veggie)
                    carbs_per_gram = highest_carb_veggie["carbohydrate(g)PerBaseGrams"] / highest_carb_veggie["baseGrams"]
                    
                    # Calculate how much we need to reduce
                    grams_to_reduce = abs(excess_carbs) / carbs_per_gram
                    
                    # Ensure we don't reduce below minimum veggie amount
                    min_veggie_grams = MIN_VEGGIES_GRAM if hasattr(self, 'MIN_VEGGIES_GRAM') else current_grams * 0.5  # fallback to 50%
                    max_reducible = current_grams - min_veggie_grams
                    grams_to_reduce = min(grams_to_reduce, max_reducible)
                    
                    if grams_to_reduce > 0:
                        scaler_reduction = grams_to_reduce / highest_carb_veggie["baseGrams"]
                        highest_carb_veggie["scaler"] = max(0.1, highest_carb_veggie["scaler"] - scaler_reduction)
                        return True
            
            return reduced
            
        # -------------------------
        # Process the different cases in a logical order
        # -------------------------
        
        # CASE 1: All three are low (calories, protein, and carbs)
        if ('kcal too low' in notes and 
            'protein(g) too low' in notes and 
            'carbohydrate(g) too low' in notes):
            
            # First, try to add veggies for carbs
            carbs_needed = get_carbs_needed()
            add_veggies_for_carbs(carbs_needed)
            
            # Recalculate nutrition
            current_nutrition = self.calculate_total_nutrition(final_recipe)
            notes = self.format_result(final_recipe, current_nutrition)['results']['notes']
            
            # If still need carbs, try starch
            if 'carbohydrate(g) too low' in notes:
                carbs_needed = get_carbs_needed()
                add_starch_for_carbs(carbs_needed)
            
            # Recalculate nutrition
            current_nutrition = self.calculate_total_nutrition(final_recipe)
            notes = self.format_result(final_recipe, current_nutrition)['results']['notes']
            
            # Finally, address protein needs
            if 'protein(g) too low' in notes:
                protein_needed = get_protein_needed()
                add_protein(protein_needed)
                
        # CASE 2: Calories and carbs are low (but protein is OK)
        elif ('kcal too low' in notes and 
             'carbohydrate(g) too low' in notes):
            
            # First, try to add veggies for carbs
            carbs_needed = get_carbs_needed()
            add_veggies_for_carbs(carbs_needed)
            
            # Recalculate nutrition
            current_nutrition = self.calculate_total_nutrition(final_recipe)
            notes = self.format_result(final_recipe, current_nutrition)['results']['notes']
            
            # If still need carbs, try starch
            if 'carbohydrate(g) too low' in notes:
                carbs_needed = get_carbs_needed()
                add_starch_for_carbs(carbs_needed)
            
            # Recalculate nutrition
            current_nutrition = self.calculate_total_nutrition(final_recipe)
            notes = self.format_result(final_recipe, current_nutrition)['results']['notes']
            
            # If calories still low, try doubling sauce
            if 'kcal too low' in notes:
                double_sauce()
        
        # CASE 3: Calories and protein are low (but carbs are OK)
        elif ('kcal too low' in notes and 
             'protein(g) too low' in notes):
            
            # Address protein needs (will also help with calories)
            protein_needed = get_protein_needed()
            add_protein(protein_needed)
            
            # Recalculate nutrition
            current_nutrition = self.calculate_total_nutrition(final_recipe)
            notes = self.format_result(final_recipe, current_nutrition)['results']['notes']
            
            # If calories still low, try doubling sauce
            if 'kcal too low' in notes:
                double_sauce()
         # CASE 9: Calories low, protein low, carbs high
        elif ('kcal too low' in notes and
              'protein(g) too low' in notes and
              'carbohydrate(g) too high' in notes):
              
            # First reduce carbs (starch/veggies) 
            carbs_needed = get_carbs_needed()  # Will be negative
            reduce_carbs(carbs_needed)
            
            # Then increase protein to address both protein and calories
            current_nutrition = self.calculate_total_nutrition(final_recipe)
            notes = self.format_result(final_recipe, current_nutrition)['results']['notes']
            
            if 'protein(g) too low' in notes:
                protein_needed = get_protein_needed()  # Will be positive
                add_protein(protein_needed)

        # CASE 8: Calories high and carbs high
        elif ('kcal too high' in notes and 
              'carbohydrate(g) too high' in notes and
              'protein(g) too high' not in notes):
            
            # Reducing carbs will help with both issues at once
            carbs_needed = get_carbs_needed()  # Will be negative when too high
            reduce_carbs(carbs_needed)
        
                        
        # CASE 7: Calories high and protein high (carbs are OK)
        elif ('kcal too high' in notes and 
              'protein(g) too high' in notes):
            
            # Reducing protein will help with both issues at once
            protein_needed = get_protein_needed()  # Will be negative when too high
            reduce_protein(protein_needed)

        
        # CASE 10: Calories high, protein low, carbs high
        elif ('kcal too high' in notes and
              'carbohydrate(g) too high' in notes):
              
            # Reduce carbs first
            carbs_needed = get_carbs_needed()  # Will be negative
            reduce_carbs(carbs_needed)
            
            # Recalculate
            current_nutrition = self.calculate_total_nutrition(final_recipe)
            notes = self.format_result(final_recipe, current_nutrition)['results']['notes']
            
        # Recalculate final nutrition
        current_nutrition = self.calculate_total_nutrition(final_recipe)
        return final_recipe, current_nutrition
   
   
   
    def solve(self, max_iterations=1000):
        """Modified solver with sequential ingredient adjustment strategy"""
        if not self.dish:
            return None, None

        recipe = self.dish['ingredients']
        dish_name = self.dish['dishName']
        best_recipe = None
        best_deviation = float('inf')
        best_nutrition = None
        
        # Check if any protein items are special protein items
        self.is_special_yogurt_protein = any(
            any(item.lower() in ing['ingredientName'].lower() for item in SPECIAL_YOGURT_PROTEIN_ITEM_KEYWORDS)
            for ing in recipe if ing['component'] == 'protein'
        )
        
        # Set the veggies limit based on dish type
        veggies_limit = MAX_SPECIAL_FRUIT_SNACK_DISH_VEGGIES_GRAM if self.is_special_fruit_snack else MAX_VEGGIES_GRAM
        # Check if the dish is a special fruit snack
        self.is_special_fruit_snack = any(keyword in dish_name.lower() for keyword in SPECIAL_FRUIT_SNACK_DISH_KEYWORDS)
        # Set up the max protein limit based on the protein type
        self.protein_max = next((MAX_PROTEIN_PER_TYPE.get(i.get('protein_type', '').lower(), 500) for i in recipe if i.get('component') == 'protein' and i.get('protein_type', '').lower() != 'ignore'), 500)
        
        # Check if the dish only contains one or two components (excluding sauce and garnish)
        unique_components = set(item['component'] for item in recipe if item['component'] not in {'sauce', 'garnish'})
        non_sauce_garnish_kcal = sum(
            item['kcalPerBaseGrams'] for item in recipe if item['component'] not in {'sauce', 'garnish'}
        )
        sauce_garnish_kcal = sum(
            item['kcalPerBaseGrams'] for item in recipe if item['component'] in {'sauce', 'garnish'}
        )

        if (len(recipe) == 1):
            for item in recipe:
                item['scaler'] = round((self.customer_requirements['kcal'] - sauce_garnish_kcal) / non_sauce_garnish_kcal)
            return self.format_result(recipe, self.calculate_total_nutrition(recipe))
        elif (len(unique_components) <= 2):
            for item in recipe:
                if item['component'] == 'sauce':
                    if self.double_sauce:
                        item['scaler'] = 2
                    else:
                        item['scaler'] = 1.0
                elif item['component'] == 'garnish':
                    item['scaler'] = 1.0
                else:
                    item['scaler'] = (self.customer_requirements['kcal'] - sauce_garnish_kcal) / non_sauce_garnish_kcal
            if self.is_special_yogurt_protein:
                recipe = self.adjust_component_within_limit(recipe, 'protein', MAX_SPECIAL_YOGURT_PROTEIN_GRAM)
                recipe = self.adjust_component_within_limit(recipe, 'veggies', MAX_SPECIAL_YOGURT_VEGGIES_GRAM)
            elif self.is_special_fruit_snack:
                recipe = self.adjust_component_within_limit(recipe, 'veggies', MAX_SPECIAL_FRUIT_SNACK_DISH_VEGGIES_GRAM)
            else:
                recipe = self.adjust_component_within_limit(recipe, 'veggies', MAX_VEGGIES_GRAM)
                recipe = self.adjust_component_within_limit(recipe, 'starch', MAX_STARCH_GRAM)
                recipe = self.adjust_component_within_limit(recipe, 'protein', self.protein_max)
            return self.format_result(recipe, self.calculate_total_nutrition(recipe))
        
        # Initialize recipe with base scalers
        initial_nutrition = self.calculate_total_nutrition(recipe)
        veggie_count = sum(1 for item in recipe if item['component'] == 'veggies')
        if initial_nutrition['kcal'] > 0:
            for ing in recipe:
                if ing['component'] == 'sauce':
                    if self.double_sauce:
                        ing['scaler'] = 2
                    else:
                        ing['scaler'] = 1.0
                elif ing['component'] == 'garnish':
                    ing['scaler'] = 1.0
                else:
                    nutrition_ratio = (
                        self.customer_requirements['protein(g)'] / initial_nutrition['protein(g)'] * 0.9 if ing['component'] == 'protein'
                        else self.customer_requirements['carbohydrate(g)'] / initial_nutrition['carbohydrate(g)'] * 0.3 if ing['component'] == 'starch'
                        else self.customer_requirements['dietaryFiber(g)'] / initial_nutrition['dietaryFiber(g)'] if ing['component'] == 'veggies'
                        else 1.0
                    )
                    calorie_ratio = self.customer_requirements['kcal'] / initial_nutrition['kcal']
                    ing['scaler'] = nutrition_ratio

        if self.is_special_yogurt_protein:
                recipe = self.adjust_component_within_limit(recipe, 'protein', MAX_SPECIAL_YOGURT_PROTEIN_GRAM)
                recipe = self.adjust_component_within_limit(recipe, 'veggies', MAX_SPECIAL_YOGURT_VEGGIES_GRAM)
        elif self.is_special_fruit_snack:
                recipe = self.adjust_component_within_limit(recipe, 'veggies', MAX_SPECIAL_FRUIT_SNACK_DISH_VEGGIES_GRAM)
        

        for iteration in range(max_iterations):
            # Enforce component limits
            recipe = self.adjust_component_within_limit(recipe, 'starch', MAX_STARCH_GRAM)
            recipe = self.adjust_component_above_minimum(recipe, 'starch', MIN_STARCH_GRAM)
            if self.is_special_yogurt_protein:
                recipe = self.adjust_component_within_limit(recipe, 'protein', MAX_SPECIAL_YOGURT_PROTEIN_GRAM)
                recipe = self.adjust_component_within_limit(recipe, 'veggies', MAX_SPECIAL_YOGURT_VEGGIES_GRAM)
            else:
                recipe = self.adjust_component_within_limit(recipe, 'protein', self.protein_max)
                recipe = self.adjust_component_within_limit(recipe, 'veggies', veggies_limit)
           
            # Get current state
            current_nutrition = self.calculate_total_nutrition(recipe)
            current_deviation = self.calculate_weighted_deviation(current_nutrition, self.customer_requirements, recipe)

            # Update best solution if current is better
            if current_deviation < best_deviation and self.check_recipe_constraints(recipe):
                best_deviation = current_deviation
                best_recipe = [{**ing} for ing in recipe]
                best_nutrition = current_nutrition.copy()

            # Check if solution meets requirements
            if self.is_within_nutrition_range(recipe, current_nutrition) and self.check_recipe_constraints(recipe):
                return self.format_result(recipe, current_nutrition)

            # Sequential ingredient adjustment
            adjusted_recipe = self._adjust_ingredients_sequentially(recipe, current_nutrition)
            
            # If no meaningful changes were made, break
            if self._recipes_are_similar(recipe, adjusted_recipe):
                break
            
            recipe = adjusted_recipe
        tmp_result = self.format_result(best_recipe or recipe, best_nutrition or current_nutrition)
        if tmp_result['results']['review_needed']:
            final_recipe, final_nutrition = self._final_adjustment(tmp_result, best_recipe or recipe, best_nutrition or current_nutrition)
            return self.format_result(final_recipe, final_nutrition)
        else:
            return self.format_result(best_recipe or recipe, best_nutrition or current_nutrition)
