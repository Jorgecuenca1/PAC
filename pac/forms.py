from django import forms
from .models import FuenteFinanciacion


class ImportarArchivoForm(forms.Form):
    archivo = forms.FileField(
        label='Seleccionar archivo Excel (.xlsx)',
        help_text='El archivo debe seguir el formato de la plantilla correspondiente.',
        widget=forms.FileInput(attrs={
            'class': 'form-control',
            'accept': '.xlsx,.xls'
        })
    )
    vigencia = forms.IntegerField(
        initial=2026,
        widget=forms.NumberInput(attrs={'class': 'form-control'}),
        label='Vigencia'
    )


class FuenteFinanciacionForm(forms.ModelForm):
    class Meta:
        model = FuenteFinanciacion
        fields = ['codigo', 'nombre', 'descripcion', 'presupuesto_asignado', 'vigencia', 'activa']
        widgets = {
            'codigo': forms.TextInput(attrs={'class': 'form-control'}),
            'nombre': forms.TextInput(attrs={'class': 'form-control'}),
            'descripcion': forms.Textarea(attrs={'class': 'form-control', 'rows': 3}),
            'presupuesto_asignado': forms.NumberInput(attrs={'class': 'form-control'}),
            'vigencia': forms.NumberInput(attrs={'class': 'form-control'}),
            'activa': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
        }
