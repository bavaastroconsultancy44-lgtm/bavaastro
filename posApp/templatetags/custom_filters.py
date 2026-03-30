from django import template

register = template.Library()

@register.filter
def get_item(dictionary, key):
    """
    Template filter to get an item from a dictionary by key
    """
    if dictionary is None:
        return None
    return dictionary.get(key)

@register.filter
def sum_values(iterable):
    """
    Template filter to sum values in an iterable
    """
    if iterable is None:
        return 0
    try:
        return sum(float(v) for v in iterable if v is not None)
    except (ValueError, TypeError):
        return 0
