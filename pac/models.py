from django.db import models
from django.contrib.auth.models import User
from decimal import Decimal

MESES = [
    'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
    'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
]

MESES_DISPLAY = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
]


class FuenteFinanciacion(models.Model):
    """Fuentes de financiacion con presupuesto asignado"""
    codigo = models.CharField(max_length=50, verbose_name='Codigo', blank=True)
    nombre = models.CharField(max_length=200, unique=True, verbose_name='Nombre')
    descripcion = models.TextField(blank=True, verbose_name='Descripcion')
    presupuesto_asignado = models.DecimalField(
        max_digits=20, decimal_places=2, default=0,
        verbose_name='Presupuesto Asignado'
    )
    vigencia = models.IntegerField(default=2026)
    activa = models.BooleanField(default=True)
    fecha_creacion = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = 'Fuente de Financiacion'
        verbose_name_plural = 'Fuentes de Financiacion'
        ordering = ['nombre']

    def __str__(self):
        return self.nombre

    def get_total_programado_ingresos(self):
        from django.db.models import Sum
        from django.db.models.functions import Coalesce
        return PACProgramado.objects.filter(
            vigencia=self.vigencia, tipo='INGRESO', fuente_financiacion=self.nombre
        ).aggregate(t=Coalesce(Sum('total'), Decimal('0')))['t']

    def get_total_programado_gastos(self):
        from django.db.models import Sum
        from django.db.models.functions import Coalesce
        return PACProgramado.objects.filter(
            vigencia=self.vigencia, tipo='GASTO', fuente_financiacion=self.nombre
        ).aggregate(t=Coalesce(Sum('total'), Decimal('0')))['t']

    def get_total_compromisos(self):
        from django.db.models import Sum
        from django.db.models.functions import Coalesce
        return PACEjecutadoCompromiso.objects.filter(
            vigencia=self.vigencia, tipo='GASTO', fuente_financiacion=self.nombre
        ).aggregate(t=Coalesce(Sum('total'), Decimal('0')))['t']

    def get_total_pagos_gastos(self):
        from django.db.models import Sum
        from django.db.models.functions import Coalesce
        return PACEjecutadoPago.objects.filter(
            vigencia=self.vigencia, tipo='GASTO', fuente_financiacion=self.nombre
        ).aggregate(t=Coalesce(Sum('total'), Decimal('0')))['t']

    def get_total_recaudo(self):
        from django.db.models import Sum
        from django.db.models.functions import Coalesce
        return PACEjecutadoPago.objects.filter(
            vigencia=self.vigencia, tipo='INGRESO', fuente_financiacion=self.nombre
        ).aggregate(t=Coalesce(Sum('total'), Decimal('0')))['t']

    def get_saldo_disponible(self):
        return self.presupuesto_asignado - self.get_total_compromisos()

    def get_porcentaje_ejecucion(self):
        if self.presupuesto_asignado:
            return round(float(self.get_total_compromisos()) / float(self.presupuesto_asignado) * 100, 1)
        return 0

    def get_porcentaje_pagos(self):
        compromisos = self.get_total_compromisos()
        if compromisos:
            return round(float(self.get_total_pagos_gastos()) / float(compromisos) * 100, 1)
        return 0


class PACBase(models.Model):
    """Modelo base para todos los modulos PAC - replica la estructura del Excel real"""
    TIPO_CHOICES = [
        ('INGRESO', 'Ingreso'),
        ('GASTO', 'Gasto'),
    ]
    CATEGORIA_CHOICES = [
        ('SALDO_INICIAL', 'Saldo Inicial'),
        ('INGRESO_CORRIENTE', 'Ingresos Corrientes'),
        ('INGRESO_CAPITAL', 'Ingresos de Capital'),
        ('FUNCIONAMIENTO', 'Funcionamiento'),
        ('INVERSION', 'Inversion'),
        ('DEUDA', 'Servicio a la Deuda'),
        ('RESERVAS', 'Reservas Presupuestales'),
        ('CUENTAS_POR_PAGAR', 'Cuentas por Pagar'),
    ]

    vigencia = models.IntegerField(default=2026)
    tipo = models.CharField(max_length=10, choices=TIPO_CHOICES)
    categoria = models.CharField(max_length=30, choices=CATEGORIA_CHOICES, blank=True, default='')
    codigo_rubro = models.CharField(max_length=200, verbose_name='Codigo Rubro')
    nombre_rubro = models.CharField(max_length=500, verbose_name='Nombre Rubro')
    fuente_financiacion = models.CharField(max_length=200, verbose_name='Fuente de Financiacion', blank=True, default='')

    # Campos de apropiacion (columnas D-I del Excel)
    apropiacion_inicial = models.DecimalField(max_digits=20, decimal_places=2, default=0, verbose_name='Aprop. Inicial')
    adiciones = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    reduccion = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    creditos = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    contracreditos = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    apropiacion_definitiva = models.DecimalField(max_digits=20, decimal_places=2, default=0, verbose_name='Aprop. Definitiva')

    # Campos mensuales (columnas J-U del Excel)
    enero = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    febrero = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    marzo = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    abril = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    mayo = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    junio = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    julio = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    agosto = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    septiembre = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    octubre = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    noviembre = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    diciembre = models.DecimalField(max_digits=20, decimal_places=2, default=0)
    total = models.DecimalField(max_digits=20, decimal_places=2, default=0)

    # Metadata
    es_subtotal = models.BooleanField(default=False)
    fila_excel = models.IntegerField(default=0, help_text='Fila original del Excel')
    fecha_carga = models.DateTimeField(auto_now_add=True)
    usuario = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)

    class Meta:
        abstract = True
        ordering = ['fila_excel', 'tipo', 'codigo_rubro']

    def __str__(self):
        return f"{self.tipo} - {self.codigo_rubro} - {self.nombre_rubro[:50]}"

    def calcular_total(self):
        self.total = sum([
            self.enero, self.febrero, self.marzo, self.abril,
            self.mayo, self.junio, self.julio, self.agosto,
            self.septiembre, self.octubre, self.noviembre, self.diciembre
        ])
        return self.total

    def save(self, *args, **kwargs):
        if not self.total:
            self.calcular_total()
        if not self.apropiacion_definitiva:
            self.apropiacion_definitiva = (
                self.apropiacion_inicial + self.adiciones - self.reduccion
                + self.creditos - self.contracreditos
            )
        super().save(*args, **kwargs)

    def get_valores_mensuales(self):
        return [
            self.enero, self.febrero, self.marzo, self.abril,
            self.mayo, self.junio, self.julio, self.agosto,
            self.septiembre, self.octubre, self.noviembre, self.diciembre
        ]


class AIMInicial(PACBase):
    """PAC 2026 AIM INICIAL - Apropiacion Inicial Modificada"""
    class Meta(PACBase.Meta):
        verbose_name = 'AIM Inicial'
        verbose_name_plural = 'AIM Iniciales'


class PACProgramado(PACBase):
    """PAC Programado mensual - Hoja PROG PAC INGRESOS-GASTOS"""
    class Meta(PACBase.Meta):
        verbose_name = 'PAC Programado'
        verbose_name_plural = 'PAC Programados'


class PACEjecutadoCompromiso(PACBase):
    """PAC Ejecutado Compromisos - Hoja PAC EJECUTADO COMPROMISOS"""
    class Meta(PACBase.Meta):
        verbose_name = 'PAC Ejecutado Compromiso'
        verbose_name_plural = 'PAC Ejecutados Compromisos'


class PACEjecutadoPago(PACBase):
    """PAC Ejecutado Pagos - Hoja PAC EJECUTADO PAGOS"""
    class Meta(PACBase.Meta):
        verbose_name = 'PAC Ejecutado Pago'
        verbose_name_plural = 'PAC Ejecutados Pagos'


class CargaArchivo(models.Model):
    """Log de cargas de archivos"""
    TIPO_CHOICES = [
        ('AIM_INICIAL', 'AIM Inicial'),
        ('PROGRAMADO', 'PAC Programado'),
        ('EJECUTADO_COMPROMISO', 'PAC Ejecutado - Compromisos'),
        ('EJECUTADO_PAGO', 'PAC Ejecutado - Pagos'),
    ]
    tipo = models.CharField(max_length=30, choices=TIPO_CHOICES)
    archivo = models.FileField(upload_to='importaciones/')
    fecha_carga = models.DateTimeField(auto_now_add=True)
    usuario = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    registros_cargados = models.IntegerField(default=0)
    observaciones = models.TextField(blank=True)

    class Meta:
        verbose_name = 'Carga de Archivo'
        verbose_name_plural = 'Cargas de Archivos'
        ordering = ['-fecha_carga']

    def __str__(self):
        return f"{self.get_tipo_display()} - {self.fecha_carga.strftime('%Y-%m-%d %H:%M')}"
