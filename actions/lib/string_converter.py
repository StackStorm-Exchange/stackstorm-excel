from ast import literal_eval


def convert_string_to_float_int(string):

    if isinstance(string, str) or isinstance(string, unicode):
        try:
            # evaluate the string to float, int or bool
            value = literal_eval(string)
        except Exception:
            value = string
    else:
        value = string

    # check if the the value is float or int
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return value
    return value
