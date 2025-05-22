def format_currency(value):
    if isinstance(value, (int, float)):
        return f"${value:,.2f}"
    return value

def validate_week_number(week):
    if not isinstance(week, int) or week < 1:
        raise ValueError("Week number must be a positive integer.")

def extract_column_names(dataframe):
    return list(dataframe.columns)

def is_valid_payment_method(payment_method):
    valid_methods = ['Check', 'Credit Card', 'ACH']
    return payment_method in valid_methods

def get_qualified_status(account_type, qualifications):
    return "Qualified" if account_type in qualifications else "Non-Qualified"