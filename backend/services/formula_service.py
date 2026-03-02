def generate_formula(data):
    headers = data.headers

    if "Sales" in headers:
        return "=SUM(A2:A100)"

    return "=A1"
