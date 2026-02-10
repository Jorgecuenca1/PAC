from django import template
from django.contrib.humanize.templatetags.humanize import intcomma

register = template.Library()


@register.filter
def formato_moneda(value):
    """Formatea un valor como moneda colombiana"""
    try:
        value = float(value)
        if value < 0:
            return f"-$ {intcomma(abs(int(value)))}"
        return f"$ {intcomma(int(value))}"
    except (ValueError, TypeError):
        return "$ 0"


@register.filter
def formato_porcentaje(value):
    """Formatea un valor como porcentaje"""
    try:
        return f"{float(value):.1f}%"
    except (ValueError, TypeError):
        return "0.0%"


@register.filter
def get_item(dictionary, key):
    """Obtiene un elemento de un diccionario por clave"""
    if isinstance(dictionary, dict):
        return dictionary.get(key)
    return None


@register.filter
def index(sequence, position):
    """Obtiene un elemento de una secuencia por posición"""
    try:
        return sequence[position]
    except (IndexError, TypeError):
        return None


@register.filter
def color_porcentaje(value):
    """Retorna una clase de color según el porcentaje"""
    try:
        val = float(value)
        if val >= 90:
            return 'text-success'
        elif val >= 60:
            return 'text-warning'
        else:
            return 'text-danger'
    except (ValueError, TypeError):
        return 'text-muted'


@register.filter
def bg_porcentaje(value):
    """Retorna un color de fondo según el porcentaje"""
    try:
        val = float(value)
        if val >= 90:
            return 'bg-success'
        elif val >= 60:
            return 'bg-warning'
        else:
            return 'bg-danger'
    except (ValueError, TypeError):
        return 'bg-secondary'
