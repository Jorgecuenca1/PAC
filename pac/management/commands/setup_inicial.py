from django.core.management.base import BaseCommand
from django.contrib.auth.models import User
from pac.models import FuenteFinanciacion


class Command(BaseCommand):
    help = 'Crea el usuario admin y las fuentes de financiacion iniciales'

    def handle(self, *args, **options):
        # Crear superusuario
        if not User.objects.filter(username='admin').exists():
            User.objects.create_superuser(
                username='admin',
                email='admin@pac.gov.co',
                password='admin123',
                first_name='Administrador',
                last_name='PAC'
            )
            self.stdout.write(self.style.SUCCESS('Usuario admin creado (admin / admin123)'))
        else:
            self.stdout.write(self.style.WARNING('El usuario admin ya existe'))

        # Crear fuentes de financiacion (basadas en PlanFinanciero)
        fuentes = [
            ('ICDE', 'ICDE', 'Ingresos Corrientes de Destinacion Especifica', 5000000000),
            ('ICLD1', 'ICLD 1', 'Ingresos Corrientes de Libre Destinacion 1', 3000000000),
            ('ICLD2', 'ICLD 2', 'Ingresos Corrientes de Libre Destinacion 2', 2000000000),
            ('ESTAM', 'ESTAMPILLAS', 'Estampillas', 500000000),
            ('SGP', 'SGP', 'Sistema General de Participaciones', 3000000000),
            ('COFIN', 'COFINANCIACION', 'Cofinanciacion', 800000000),
            ('CRED', 'CREDITO', 'Credito', 600000000),
            ('REGAL', 'REGALIAS', 'Regalias', 1500000000),
            ('ICDES', 'ICDE SALUD', 'ICDE Salud', 1000000000),
        ]

        created_count = 0
        for codigo, nombre, desc, presupuesto in fuentes:
            _, created = FuenteFinanciacion.objects.get_or_create(
                nombre=nombre,
                defaults={
                    'codigo': codigo,
                    'descripcion': desc,
                    'presupuesto_asignado': presupuesto,
                    'vigencia': 2026,
                }
            )
            if created:
                created_count += 1

        self.stdout.write(self.style.SUCCESS(f'{created_count} fuentes de financiacion creadas'))
