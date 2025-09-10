import io
import re
from flask import Flask, request, jsonify, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule, DataBarRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.cell import MergedCell

# --- Inicialización de la Aplicación Flask ---
app = Flask(__name__)

# --- Funciones de Ayuda (sin cambios) ---

def apply_styles_to_cell(cell, style_data):
    """Aplica estilos a una celda a partir de un diccionario de configuración."""
    if not isinstance(style_data, dict): return
    if 'font' in style_data: cell.font = Font(**style_data['font'])
    if 'fill' in style_data:
        if 'pattern' in style_data['fill'] and 'fill_type' not in style_data['fill']:
            style_data['fill']['fill_type'] = style_data['fill'].pop('pattern')
        cell.fill = PatternFill(**style_data['fill'])
    if 'border' in style_data:
        border_styles = style_data['border']
        cell.border = Border(left=Side(**border_styles.get('left', {})), right=Side(**border_styles.get('right', {})), top=Side(**border_styles.get('top', {})), bottom=Side(**border_styles.get('bottom', {})))
    if 'alignment' in style_data: cell.alignment = Alignment(**style_data['alignment'])
    if 'numFmt' in style_data: cell.number_format = style_data['numFmt']

def create_chart_from_spec(worksheet, chart_spec):
    """Crea y añade un gráfico a la hoja de cálculo según la especificación."""
    chart_type = chart_spec.get('type', 'bar').lower()
    if chart_type in ['bar', 'col']:
        chart = BarChart(); chart.type = chart_type
    elif chart_type == 'line': chart = LineChart()
    elif chart_type == 'pie': chart = PieChart()
    else: return

    chart.title = chart_spec.get('title', 'Gráfico sin Título')
    chart.style = chart_spec.get('style', 10)
    if 'y_axis_title' in chart_spec: chart.y_axis.title = chart_spec['y_axis_title']
    if 'x_axis_title' in chart_spec: chart.x_axis.title = chart_spec['x_axis_title']

    data_range_str = chart_spec['data_range']
    if '!' in data_range_str:
        data_range_str = data_range_str.split('!')[-1]
    
    category_range_str = chart_spec['category_range']
    if '!' in category_range_str:
        category_range_str = category_range_str.split('!')[-1]

    data = Reference(worksheet, range_string=data_range_str)
    cats = Reference(worksheet, range_string=category_range_str)
    
    chart.add_data(data, titles_from_data=chart_spec.get('titles_from_data', True))
    chart.set_categories(cats)
    worksheet.add_chart(chart, chart_spec.get('position', 'E1'))

# --- Endpoint Principal ---

@app.route('/create-excel', methods=['POST'])
def create_excel():
    """Endpoint para crear un archivo Excel a partir de datos JSON."""
    try:
        json_data = request.get_json()
        if not json_data:
            return jsonify({"error": "El cuerpo de la solicitud no contiene un JSON válido."}), 400

        analysis_data = json_data.get('analysisData', [])
        conditional_rules = json_data.get('conditionalFormattingRules', [])
        chart_specs = json_data.get('charts', [])
        merge_cells_list = json_data.get('mergeCells', [])

        wb = Workbook()
        ws = wb.active
        ws.title = "Reporte Generado"

        # 1. Escribir datos y aplicar estilos
        for row_details in analysis_data:
            for cell_data in row_details:
                if isinstance(cell_data, dict) and 'address' in cell_data:
                    cell = ws[cell_data['address']]
                    cell.value = cell_data.get('value')
                    apply_styles_to_cell(cell, cell_data.get('style'))

        # 2. Combinar celdas
        # --- INICIO DE LA NUEVA CORRECCIÓN ---
        for cell_range in merge_cells_list:
            # Se asegura de que el rango no contenga el nombre de la hoja, que causa el error.
            if '!' in cell_range:
                cell_range = cell_range.split('!')[-1]
            ws.merge_cells(cell_range)
        # --- FIN DE LA NUEVA CORRECCIÓN ---

        # 3. Aplicar formato condicional
        for rule_info in conditional_rules:
            style = rule_info.get('style', {})
            dxf = DifferentialStyle(font=Font(**style.get('font', {})), fill=PatternFill(**style.get('fill', {})))
            rule_type = rule_info.get('type')
            if rule_type == 'dataBar':
                rule = DataBarRule(start_type='min', end_type='max', color=rule_info.get("color", "638EC6"))
            else:
                rule_params = {'type': rule_type, 'dxf': dxf}
                if 'operator' in rule_info: rule_params['operator'] = rule_info['operator']
                if rule_type == 'containsText' and 'formulae' in rule_info:
                    rule_params['text'] = rule_info['formulae'][0]
                elif 'formulae' in rule_info:
                    rule_params['formula'] = rule_info['formulae']
                rule = Rule(**rule_params)
            ws.conditional_formatting.add(rule_info['ref'], rule)

        # 4. Crear gráficos
        for spec in chart_specs:
            create_chart_from_spec(ws, spec)

        # 5. Ajustar ancho de columnas
        column_widths = {}
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell): continue
                length = len(str(cell.value)) if cell.value else 0
                if cell.column not in column_widths or length > column_widths[cell.column]:
                    column_widths[cell.column] = length
        for col_idx, max_length in column_widths.items():
            column_letter = get_column_letter(col_idx)
            adjusted_width = min(max_length + 2, 50) 
            ws.column_dimensions[column_letter].width = adjusted_width

        # 6. Guardar en buffer de memoria
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        # 7. Enviar el archivo
        return send_file(
            buffer,
            as_attachment=True,
            download_name='reporte_final_corregido.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        # Imprime el error en la consola del servidor para una mejor depuración
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Error interno del servidor: {str(e)}"}), 500

# --- Bloque de Ejecución ---
if __name__ == '__main__':
    app.run(debug=True, port=5000)
