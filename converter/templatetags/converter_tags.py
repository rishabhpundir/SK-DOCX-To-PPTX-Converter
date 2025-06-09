# Create these directories first:
# converter/templatetags/
# Then add this file: converter/templatetags/converter_tags.py

from django import template

register = template.Library()


@register.filter
def get_item(dictionary, key):
    """Template filter to get dictionary item by key"""
    return dictionary.get(key, '')


@register.filter
def file_size(file_field):
    """Return human-readable file size"""
    try:
        size = file_field.size
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"
    except:
        return "Unknown"


# Also create: converter/templatetags/__init__.py (empty file)