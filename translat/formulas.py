import formulas

def is_valid_excel_formula(expr: str) -> bool:
    try:
        formulas.Parser().ast(expr)
        return True
    except Exception:
        return False

