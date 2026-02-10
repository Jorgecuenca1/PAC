from django.urls import path
from . import views

urlpatterns = [
    # Dashboard
    path('', views.dashboard, name='dashboard'),

    # AIM Inicial
    path('aim-inicial/', views.aim_inicial, name='aim_inicial'),
    path('aim-inicial/importar/', views.importar_aim_inicial, name='importar_aim_inicial'),

    # PAC Programado
    path('pac-programado/', views.pac_programado, name='pac_programado'),
    path('pac-programado/importar/', views.importar_pac_programado, name='importar_pac_programado'),

    # PAC Ejecutado Compromisos
    path('pac-ejecutado-compromisos/', views.pac_ejecutado_compromisos, name='pac_ejecutado_compromisos'),
    path('pac-ejecutado-compromisos/importar/', views.importar_pac_compromisos, name='importar_pac_compromisos'),

    # PAC Ejecutado Pagos
    path('pac-ejecutado-pagos/', views.pac_ejecutado_pagos, name='pac_ejecutado_pagos'),
    path('pac-ejecutado-pagos/importar/', views.importar_pac_pagos, name='importar_pac_pagos'),

    # Seguimiento
    path('seguimiento/ingresos/', views.seguimiento_ingresos, name='seguimiento_ingresos'),
    path('seguimiento/gastos/', views.seguimiento_gastos, name='seguimiento_gastos'),
    path('seguimiento/compromisos-vs-pagos/', views.seguimiento_compromisos_vs_pagos,
         name='seguimiento_compromisos_vs_pagos'),

    # Fuentes de Financiación
    path('fuentes/', views.fuentes_financiacion, name='fuentes_financiacion'),
    path('fuentes/crear/', views.fuente_crear, name='fuente_crear'),
    path('fuentes/<int:pk>/editar/', views.fuente_editar, name='fuente_editar'),
    path('fuentes/<int:pk>/eliminar/', views.fuente_eliminar, name='fuente_eliminar'),
    path('fuentes/<int:pk>/detalle/', views.fuente_detalle, name='fuente_detalle'),

    # Reportes
    path('reportes/', views.reportes, name='reportes'),

    # Exportar Excel
    path('exportar/seguimiento/<str:tipo>/', views.exportar_seguimiento_excel, name='exportar_seguimiento'),
    path('exportar/reporte-fuentes/', views.exportar_reporte_fuentes_excel, name='exportar_reporte_fuentes'),

    # Plantillas de ejemplo
    path('plantilla/<str:tipo>/', views.descargar_plantilla, name='descargar_plantilla'),

    # Eliminar datos
    path('eliminar/<str:tipo>/', views.eliminar_datos, name='eliminar_datos'),
]
