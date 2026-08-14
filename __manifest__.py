{
    'name': 'L10n Ar Import Arca Excel - Importación de Mis Comprobantes',
    'version': '18.0.1.9.0',
    'category': 'Accounting',
    'summary': 'Importe masivamente facturas de compra y venta desde Excel/CSV de ARCA. Automatice la carga de datos y evite errores manuales.',
    'description': """
L10n Ar Import Arca Excel
=========================

Importe masivamente facturas de compra y venta desde Excel/CSV de ARCA. Automatice la carga de datos y evite errores manuales.

Cargar facturas manualmente es propenso a errores y consume mucho tiempo. Este módulo permite importar directamente el archivo de "Mis Comprobantes" (Excel) o el "Libro de Compras" (CSV) de ARCA, creando automáticamente facturas, clientes/proveedores y validando datos clave.

Características Principales
---------------------------
*   **Doble Origen de Datos:** Detecta automáticamente si el archivo subido es el Excel de "Mis Comprobantes" o el CSV del "Libro de Compras" (este último solo para Compras).
*   **Importación Inteligente:** Detecta automáticamente si son Compras o Ventas, y asigna el diario correcto en ventas según el punto de venta.
*   **Validación de Datos:** Verifica CUITs, impuestos y evita duplicados antes de confirmar la importación.
*   **Gestión de Partners:** Busca partners por CUIT y los crea automáticamente si no existen en el sistema.
*   **Historial de Auditoría:** Mantiene un registro detallado de todas las importaciones realizadas para facilitar el control.
*   **Integración CIE:** Transcribe automáticamente el Código de Autorización Electrónico (CAE) a las facturas (no disponible al importar desde "Libro de Compras", que no incluye ese dato).

Uso
---
1.  Descargue el archivo Excel de "Mis Comprobantes" o el CSV del "Libro de Compras" desde el portal de ARCA.
2.  Vaya a **Contabilidad > Importar de ARCA > Nuevo Importación**.
3.  Suba el archivo y seleccione el tipo de operación (para "Libro de Compras" siempre es Compras).
4.  Haga clic en **Analizar Excel**, revise la vista previa y si todo está correcto, confirme.

Autor
-----
*   Crumges

Mantenedor
----------
Este módulo es mantenido por Crumges.
""",
    'author': 'Crumges',
    'website': 'https://crumges.com',
    'license': 'LGPL-3',
    'depends': [
        'base',
        'account',
        'mail',
        'l10n_ar',
    ],
    'external_dependencies': {'python': ['openpyxl']},
    'data': [
        'security/ir.model.access.csv',
        'views/import_history_view.xml',
        'views/import_wizard_view.xml',
        'data/menu_item.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'l10n_ar_import_arca_excel/static/src/css/arca_wizard.css',
        ],
    },
    'images': ['static/description/icon.png'],
    'installable': True,
    'application': False,
    'maintainers': ['Crumges'],
}
