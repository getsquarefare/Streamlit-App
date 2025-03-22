import numpy as np
import random
import math

SPECIAL_YOGURT_PROTEIN_ITEM_KEYWORDS = ["overnight oats", "yogurt", "yoghurt"]
MAX_SPECIAL_YOGURT_PROTEIN_GRAM = 400
MAX_SPECIAL_YOGURT_PROTEIN_ITEM_TOTAL_MEAL = 500
SPECIAL_FRUIT_SNACK_DISH_KEYWORDS = ['seasonal fruit salad']
MAX_SPECIAL_FRUIT_SNACK_DISH_VEGGIES_GRAM = 220
MAX_VEGGIES_GRAM = 300
MAX_STARCH_GRAM = 280
MIN_STARCH_GRAM = 50
MAX_PROTEIN_PER_TYPE = {"meat": 200, "fish": 220, "tofu": 350, "vegan": 200}

class NewDishOptimizer:
    def __init__(self, grouped_ingredients, customer_requirements, nutrients, nutrient_constraints,
                 garnish_grams=None, sauce_grams=None, veggie_ge_starch=True, 
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
        self.sauce_grams = sauce_grams
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
            if self.sauce_grams is not None:
                sauce_total = sum(self.get_effective_grams(i) for i in recipe if i['component'] == 'sauce')
                if sauce_total != self.sauce_grams:
                    # print(f"Sauce constraint violation: {sauce_total:.1f}g vs required {self.sauce_grams}g")
                    return False

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
                    "ingredientName": ing["ingredientName"],
                    "Component": ing["component"],
                    "Grams": round(ing["baseGrams"] * ing.get("scaler"), 2),
                    "ingredientId": ing["ingredientId"]
                }
                for ing in recipe
            ]
        }
        
        # Format explanation string
        explanation = (
            f"\nAchieved / Target Nutrition:\n"
            f"kcal: {nutrition_totals['kcal']:.1f} / {self.customer_requirements['kcal']:.1f} "
            f"({(nutrition_totals['kcal']/self.customer_requirements['kcal']*100 if self.customer_requirements['kcal'] != 0 else 0):.1f}% of target)\n"
            f"protein(g): {nutrition_totals['protein(g)']:.1f} / {self.customer_requirements['protein(g)']:.1f} "
            f"({(nutrition_totals['protein(g)']/self.customer_requirements['protein(g)']*100 if self.customer_requirements['protein(g)'] != 0 else 0):.1f}% of target)\n"
            f"fat(g): {nutrition_totals['fat(g)']:.1f} / {self.customer_requirements['fat(g)']:.1f} "
            f"({(nutrition_totals['fat(g)']/self.customer_requirements['fat(g)']*100 if self.customer_requirements['fat(g)'] != 0 else 0):.1f}% of target)\n"
            f"dietaryFiber(g): {nutrition_totals['dietaryFiber(g)']:.1f} / {self.customer_requirements['dietaryFiber(g)']:.1f} "
            f"({(nutrition_totals['dietaryFiber(g)']/self.customer_requirements['dietaryFiber(g)']*100 if self.customer_requirements['dietaryFiber(g)'] != 0 else 0):.1f}% of target)\n"
            f"carbohydrate(g): {nutrition_totals['carbohydrate(g)']:.1f} / {self.customer_requirements['carbohydrate(g)']:.1f} "
            f"({(nutrition_totals['carbohydrate(g)']/self.customer_requirements['carbohydrate(g)']*100 if self.customer_requirements['carbohydrate(g)'] != 0 else 0):.1f}% of target)"
        )
        
        # Format updated nutrition info
        updated_nutrition_info = {
            "Calories": round(nutrition_totals['kcal'], 1),
            "Protein": round(nutrition_totals['protein(g)'], 1),
            "Carbohydrates": round(nutrition_totals['carbohydrate(g)'], 1),
            "Fiber": round(nutrition_totals['dietaryFiber(g)'], 1),
            "Fat": round(nutrition_totals['fat(g)'], 1)
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
            if potential_grams > MAX_VEGGIES_GRAM:
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
                    new_scaler = current_scaler * (1.0 + adjustment)
                    
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
        
        # Case 1: Need more carbs but protein is sufficient
        if ('carbohydrate(g) too low' in notes and 
            'protein(g) too low' not in notes):
            
            # Try veggies first
            total_veggie_grams = sum(self.get_effective_grams(item) for item in final_recipe if item["component"] == "veggies")
            remaining_veggie_capacity = MAX_VEGGIES_GRAM - total_veggie_grams
            
            if remaining_veggie_capacity > 0:
                # Find veggie with highest carb content
                highest_carb_veggie = max(
                    (item for item in final_recipe if item["component"] == "veggies"), 
                    key=lambda x: x["carbohydrate(g)PerBaseGrams"]
                )
                
                # Calculate how many carbs we can add through veggies
                carbs_per_gram = highest_carb_veggie["carbohydrate(g)PerBaseGrams"] / highest_carb_veggie["baseGrams"]
                carbs_needed = float(notes.split("carbohydrate(g) too low: ")[1].split(" < ")[1]) - float(notes.split("carbohydrate(g) too low: ")[1].split(" < ")[0])
                grams_to_add = min(remaining_veggie_capacity, carbs_needed / carbs_per_gram)
                highest_carb_veggie["scaler"] += grams_to_add / highest_carb_veggie["baseGrams"]
                
                # Update remaining carbs needed
                current_nutrition = self.calculate_total_nutrition(final_recipe)
                notes = self.format_result(final_recipe, current_nutrition)['results']['notes']
            
            # If still need carbs, try starch
            if 'carbohydrate(g) too low' in notes:
                starch_items = [item for item in final_recipe if item["component"] == "starch"]
                if starch_items:
                    starch_item = starch_items[0]
                    current_starch_grams = sum(self.get_effective_grams(item) for item in final_recipe 
                                            if item["component"] == "starch")
                    remaining_starch_capacity = MAX_STARCH_GRAM - current_starch_grams
                    
                    if remaining_starch_capacity > 0:
                        carbs_needed = float(notes.split("carbohydrate(g) too low: ")[1].split(" < ")[1]) - float(notes.split("carbohydrate(g) too low: ")[1].split(" < ")[0])
                        carbs_per_gram = starch_item["carbohydrate(g)PerBaseGrams"] / starch_item["baseGrams"]
                        grams_to_add = min(remaining_starch_capacity, carbs_needed / carbs_per_gram)
                        starch_item["scaler"] += grams_to_add / starch_item["baseGrams"]
        
        # Case 2: Calories not hit but protein and carbs are sufficient
        elif ('kcal too low' in notes and 
            'protein(g) too low' not in notes and 
            'carbohydrate(g) too low' not in notes):
            
            # Try doubling sauce first - only use integer scalers (1 or 2)
            sauce_items = [item for item in final_recipe if item["component"] == "sauce"]
            if sauce_items:
                for sauce_item in sauce_items:
                    current_scaler = int(sauce_item["scaler"])  # Ensure current scaler is integer
                    if current_scaler == 1:  # Only double if currently at 1
                        sauce_item["scaler"] = 2
            
            # Recalculate nutrition after sauce adjustment
            current_nutrition = self.calculate_total_nutrition(final_recipe)
            notes = self.format_result(final_recipe, current_nutrition)['results']['notes']
            
            # If still need calories after doubling sauce, carefully increase protein
            if 'kcal too low' in notes:
                protein_items = [item for item in final_recipe if item["component"] == "protein"]
                if protein_items:
                    protein_item = protein_items[0]
                    kcal_needed = float(notes.split("kcal too low: ")[1].split(" < ")[1]) - float(notes.split("kcal too low: ")[1].split(" < ")[0])
                    kcal_per_gram = protein_item["kcalPerBaseGrams"] / protein_item["baseGrams"]
                    grams_to_add = kcal_needed / kcal_per_gram
                    
                    # Check if it's a special protein item
                    if (any(item.lower() in protein_item["ingredientName"].lower() 
                        for item in SPECIAL_YOGURT_PROTEIN_ITEM_KEYWORDS)):
                        current_grams = self.get_effective_grams(protein_item)
                        max_additional_grams = max(0, MAX_SPECIAL_YOGURT_PROTEIN_GRAM - current_grams)
                        grams_to_add = min(grams_to_add, max_additional_grams)
                    
                    protein_item["scaler"] += grams_to_add / protein_item["baseGrams"]
        
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
        self.is_special_fruit_snack = any(keyword in dish_name.lower() for keyword in SPECIAL_FRUIT_SNACK_DISH_KEYWORDS)
        veggies_limit = MAX_SPECIAL_FRUIT_SNACK_DISH_VEGGIES_GRAM if self.is_special_fruit_snack else MAX_VEGGIES_GRAM
        protein_max = next((MAX_PROTEIN_PER_TYPE.get(i.get('protein_type', '').lower(), 500) for i in recipe if i.get('component') == 'protein' and i.get('protein_type', '').lower() != 'ignore'), 500)
        
        unique_components = set(item['component'] for item in recipe if item['component'] not in {'sauce', 'garnish'})
        non_sauce_garnish_kcal = sum(
            item['kcalPerBaseGrams'] for item in recipe if item['component'] not in {'sauce', 'garnish'}
        )
        if (len(unique_components) <= 2):
            for item in recipe:
                if item['component'] in {'sauce', 'garnish'}:
                    item['scaler'] = 1
                else:
                    item['scaler'] = self.customer_requirements['kcal'] / non_sauce_garnish_kcal * (1- (len(unique_components) - 1) * 0.5)
            recipe = self.adjust_component_within_limit(recipe, 'veggies', MAX_VEGGIES_GRAM)
            recipe = self.adjust_component_within_limit(recipe, 'starch', MAX_STARCH_GRAM)
            if self.is_special_yogurt_protein:
                recipe = self.adjust_component_within_limit(recipe, 'protein', MAX_SPECIAL_YOGURT_PROTEIN_GRAM)
            else:
                recipe = self.adjust_component_within_limit(recipe, 'protein', protein_max)
            return self.format_result(recipe, self.calculate_total_nutrition(recipe))
        
        # Initialize recipe with base scalers
        initial_nutrition = self.calculate_total_nutrition(recipe)
        veggie_count = sum(1 for item in recipe if item['component'] == 'veggies')
        if initial_nutrition['kcal'] > 0:
            for ing in recipe:
                if ing['component'] == 'sauce' and self.sauce_grams:
                    ing['scaler'] = self.sauce_grams / ing['baseGrams'] if ing['baseGrams'] else 1.0
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

        

        for iteration in range(max_iterations):
            # Enforce component limits
            recipe = self.adjust_component_within_limit(recipe, 'veggies', veggies_limit)
            recipe = self.adjust_component_within_limit(recipe, 'starch', MAX_STARCH_GRAM)
            recipe = self.adjust_component_above_minimum(recipe, 'starch', MIN_STARCH_GRAM)
            if self.is_special_yogurt_protein:
                recipe = self.adjust_component_within_limit(recipe, 'protein', MAX_SPECIAL_YOGURT_PROTEIN_GRAM)
            else:
                recipe = self.adjust_component_within_limit(recipe, 'protein', protein_max)
           
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
        
# Example Usage and Testing
if __name__ == "__main__":
    # Test data setup
    grouped_ingredients = {'veggies': {'Kcal': 83.04, 'Carbohydrate, total (g)': 15.188, 'Protein (g)': 6.409999999999999, 'Fat, Total (g)': 1.3438999999999999, 'Dietary Fiber (g)': 6.703, 'Grams': 300.0}, 'garnish': {'Kcal': 15.875238095238098, 'Carbohydrate, total (g)': 0.9524761904761906, 'Protein (g)': 0.667647619047619, 'Fat, Total (g)': 1.1679047619047618, 'Dietary Fiber (g)': 0.46209523809523806, 'Grams': 10.0}, 'protein': {'Kcal': 110.96774193548387, 'Carbohydrate, total (g)': 0.6451612903225806, 'Protein (g)': 10.32258064516129, 'Fat, Total (g)': 7.096774193548387, 'Dietary Fiber (g)': 0.6451612903225806, 'Grams': 100.0}, 'sauce': {'Kcal': 47.65, 'Carbohydrate, total (g)': 0.27, 'Protein (g)': 0.07, 'Fat, Total (g)': 5.0, 'Dietary Fiber (g)': 0.0, 'Grams': 20.0}, 'starch': {'Kcal': 112.0, 'Carbohydrate, total (g)': 23.5, 'Protein (g)': 2.32, 'Fat, Total (g)': 0.83, 'Dietary Fiber (g)': 1.8, 'Grams': 100.0}}
    
    customer_requirements = {'identifier': 'Brad Riew | Lunch | riew.brad@gmail.com', 'First_Name': 'Brad', 'Last_Name': 'Riew', 'goal_calories': 700, 'goal_carbs(g)': 70, 'goal_fiber(g)': 11, 'goal_fat(g)': 23, 'goal_protein(g)': 53, 'Portion Algo Constraints': ['recpZqUPjx32nuwjZ'], 'Meal': 'Lunch', '# of snacks per day': 1, 'Kcal': 700, 'Carbohydrate, total (g)': 70, 'Protein (g)': 53, 'Fat, Total (g)': 23, 'Dietary Fiber (g)': 11}

    nutrients = ['kcal', 'protein(g)', 'fat(g)', 'dietaryFiber(g)', 'carbohydrate(g)']

    nutrient_constraints = {'Protein (g)': {'lb': 0.8, 'ub': 1.05}, 'Kcal': {'lb': 0.95, 'ub': 1.05}, 'Fat, Total (g)': {'lb': 0.5, 'ub': 1.5}, 'Dietary Fiber (g)': {'lb': 0.6, 'ub': None}, 'Carbohydrate, total (g)': {'lb': 0.5, 'ub': 1.1}}

    dish = {'dishName': 'Shroom Bowl', 'ingredients': [{'component': 'veggies', 'ingredientName': 'Shiitake Mushrooms', 'baseGrams': 20.0, 'kcalPerBaseGrams': 7.8, 'protein(g)PerBaseGrams': 0.69, 'fat(g)PerBaseGrams': 0.07, 'dietaryFiber(g)PerBaseGrams': 0.72, 'carbohydrate(g)PerBaseGrams': 1.54, 'scaler': 1.0}, {'component': 'garnish', 'ingredientName': 'Sesame Seeds', 'baseGrams': 2.0, 'kcalPerBaseGrams': 11.46, 'protein(g)PerBaseGrams': 0.35, 'fat(g)PerBaseGrams': 0.99, 'dietaryFiber(g)PerBaseGrams': 0.24, 'carbohydrate(g)PerBaseGrams': 0.47, 'scaler': 1.0}, {'component': 'garnish', 'ingredientName': 'Kimchi', 'baseGrams': 2.0, 'kcalPerBaseGrams': 0.3, 'protein(g)PerBaseGrams': 0.02, 'fat(g)PerBaseGrams': 0.01, 'dietaryFiber(g)PerBaseGrams': 0.03, 'carbohydrate(g)PerBaseGrams': 0.05, 'scaler': 1.0}, {'component': 'protein', 'ingredientName': 'Roasted Tofu', 'baseGrams': 100.0, 'kcalPerBaseGrams': 110.97, 'protein(g)PerBaseGrams': 10.32, 'fat(g)PerBaseGrams': 7.1, 'dietaryFiber(g)PerBaseGrams': 0.65, 'carbohydrate(g)PerBaseGrams': 0.65, 'scaler': 1.0}, {'component': 'veggies', 'ingredientName': 'Portobello Mushroom', 'baseGrams': 20.0, 'kcalPerBaseGrams': 6.48, 'protein(g)PerBaseGrams': 0.55, 'fat(g)PerBaseGrams': 0.06, 'dietaryFiber(g)PerBaseGrams': 0.38, 'carbohydrate(g)PerBaseGrams': 0.93, 'scaler': 1.0}, {'component': 'veggies', 'ingredientName': 'Broccoli', 'baseGrams': 100.0, 'kcalPerBaseGrams': 35.0, 'protein(g)PerBaseGrams': 2.38, 'fat(g)PerBaseGrams': 0.41, 'dietaryFiber(g)PerBaseGrams': 3.3, 'carbohydrate(g)PerBaseGrams': 7.18, 'scaler': 1.0}, {'component': 'sauce', 'ingredientName': 'Sesame Ginger Dressing', 'baseGrams': 20.0, 'kcalPerBaseGrams': 47.65, 'protein(g)PerBaseGrams': 0.07, 'fat(g)PerBaseGrams': 5.0, 'dietaryFiber(g)PerBaseGrams': 0.0, 'carbohydrate(g)PerBaseGrams': 0.27, 'scaler': 1.0}, {'component': 'garnish', 'ingredientName': 'Pickled Red Onions', 'baseGrams': 2.0, 'kcalPerBaseGrams': 0.86, 'protein(g)PerBaseGrams': 0.02, 'fat(g)PerBaseGrams': 0.0, 'dietaryFiber(g)PerBaseGrams': 0.04, 'carbohydrate(g)PerBaseGrams': 0.19, 'scaler': 1.0}, {'component': 'garnish', 'ingredientName': 'Edamame', 'baseGrams': 2.0, 'kcalPerBaseGrams': 2.8, 'protein(g)PerBaseGrams': 0.23, 'fat(g)PerBaseGrams': 0.15, 'dietaryFiber(g)PerBaseGrams': 0.1, 'carbohydrate(g)PerBaseGrams': 0.17, 'scaler': 1.0}, {'component': 'garnish', 'ingredientName': 'Cilantro', 'baseGrams': 2.0, 'kcalPerBaseGrams': 0.46, 'protein(g)PerBaseGrams': 0.04, 'fat(g)PerBaseGrams': 0.01, 'dietaryFiber(g)PerBaseGrams': 0.06, 'carbohydrate(g)PerBaseGrams': 0.07, 'scaler': 1.0}, {'component': 'veggies', 'ingredientName': 'Kale', 'baseGrams': 50.0, 'kcalPerBaseGrams': 18.0, 'protein(g)PerBaseGrams': 1.47, 'fat(g)PerBaseGrams': 0.6, 'dietaryFiber(g)PerBaseGrams': 2.0, 'carbohydrate(g)PerBaseGrams': 2.65, 'scaler': 1.0}, {'component': 'veggies', 'ingredientName': 'Cabbage', 'baseGrams': 100.0, 'kcalPerBaseGrams': 12.0, 'protein(g)PerBaseGrams': 1.1, 'fat(g)PerBaseGrams': 0.17, 'dietaryFiber(g)PerBaseGrams': 0.0, 'carbohydrate(g)PerBaseGrams': 2.23, 'scaler': 1.0}, {'component': 'starch', 'ingredientName': 'Brown Rice', 'baseGrams': 100.0, 'kcalPerBaseGrams': 112.0, 'protein(g)PerBaseGrams': 2.32, 'fat(g)PerBaseGrams': 0.83, 'dietaryFiber(g)PerBaseGrams': 1.8, 'carbohydrate(g)PerBaseGrams': 23.5, 'scaler': 1.0}, {'component': 'veggies', 'ingredientName': 'Maitake Mushrooms', 'baseGrams': 10.0, 'kcalPerBaseGrams': 3.76, 'protein(g)PerBaseGrams': 0.22, 'fat(g)PerBaseGrams': 0.03, 'dietaryFiber(g)PerBaseGrams': 0.31, 'carbohydrate(g)PerBaseGrams': 0.66, 'scaler': 1.0}]}

    # Initialize optimizer with specified constraints
    optimizer = NewDishOptimizer(
        grouped_ingredients,  # Not needed for this test
        customer_requirements=customer_requirements,
        nutrients=nutrients,
        nutrient_constraints=nutrient_constraints,
        veggie_ge_starch=True,  # Veggies must exceed starch
        min_meat_per_100_cal=20,  # 20g meat per 100 calories
        max_meal_grams_per_100_cal=200,  # Maximum 200g per 100 calories
        dish=dish
    )

    # Run optimization
    optimized_recipe = optimizer.solve()

    # Print results
    print(optimized_recipe)
    