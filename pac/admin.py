from django.contrib import admin
from .models import AIMInicial, PACProgramado, PACEjecutadoCompromiso, PACEjecutadoPago, CargaArchivo, FuenteFinanciacion


@admin.register(FuenteFinanciacion)
class FuenteFinanciacionAdmin(admin.ModelAdmin):
    list_display = ['codigo', 'nombre', 'presupuesto_asignado', 'vigencia', 'activa']
    list_filter = ['vigencia', 'activa']
    search_fields = ['nombre', 'codigo']


class PACBaseAdmin(admin.ModelAdmin):
    list_display = ['vigencia', 'tipo', 'categoria', 'codigo_rubro', 'nombre_rubro',
                    'apropiacion_definitiva', 'total', 'es_subtotal']
    list_filter = ['vigencia', 'tipo', 'categoria', 'es_subtotal']
    search_fields = ['codigo_rubro', 'nombre_rubro']


@admin.register(AIMInicial)
class AIMInicialAdmin(PACBaseAdmin):
    pass


@admin.register(PACProgramado)
class PACProgramadoAdmin(PACBaseAdmin):
    pass


@admin.register(PACEjecutadoCompromiso)
class PACEjecutadoCompromisoAdmin(PACBaseAdmin):
    pass


@admin.register(PACEjecutadoPago)
class PACEjecutadoPagoAdmin(PACBaseAdmin):
    pass


@admin.register(CargaArchivo)
class CargaArchivoAdmin(admin.ModelAdmin):
    list_display = ['tipo', 'fecha_carga', 'usuario', 'registros_cargados']
    list_filter = ['tipo', 'fecha_carga']
