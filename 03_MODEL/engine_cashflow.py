import pandas as pd

# ================================
# PARAMETERS (HARDCODE TEMPORAL)
# ================================
OCCUPANCY_RATE = 0.90
SALE_DISCOUNT = 0.90

# ================================
# MOCK INPUT (REEMPLAZAR DESPUES)
# ================================
monthly_rent_base = 1000000  # reemplazar desde extractor real

# ================================
# CORE CALCULATION
# ================================
def compute_annual_income(monthly):
    return monthly * 12 * OCCUPANCY_RATE

def compute_cashflow(monthly):
    annual = compute_annual_income(monthly)
    return {
        "monthly": monthly,
        "annual": annual
    }

def project_5_years(monthly):
    results = []
    for year in range(1, 6):
        annual = compute_annual_income(monthly)
        results.append({
            "year": year,
            "annual_income": annual
        })
    return pd.DataFrame(results)

def apply_sale(value):
    return value * SALE_DISCOUNT

# ================================
# EXECUTION
# ================================
if __name__ == "__main__":
    base = compute_cashflow(monthly_rent_base)
    projection = project_5_years(monthly_rent_base)

    sale_value = apply_sale(base["annual"])

    print("=== BASE ===")
    print(base)

    print("\n=== PROJECTION ===")
    print(projection)

    print("\n=== SALE ===")
    print(sale_value)
