import io
from flask import Flask, request, jsonify, send_file
from openpyxl import Workbook
# --- LÍNEAS DE IMPORTACIÓN CORREGIDAS ---
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.styles.differential import DifferentialStyle
# --- FIN DE LA CORRECCIÓN ---
from openpyxl.formatting.rule import Rule, DataBarRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference

# --- Inicialización de la Aplicación Flask ---
app = Flask(__name__)

# --- Funciones de Ayuda para Estilos y Formato ---

def apply_styles_to_cell(cell, style_data):
    """Aplica los estilos del JSON a una celda de OpenPyXL."""
    if not style_data or not isinstance(style_data, dict):
        return

    if 'font' in style_data:
        cell.font = Font(**style_data['font'])
    if 'fill' in style_data:
        if 'pattern' in style_data['fill']:
            style_data['fill']['fill_type'] = style_data['fill'].pop('pattern')
        cell.fill = PatternFill(**style_data['fill'])
    if 'border' in style_data:
        cell.border = Border(
            left=Side(**style_data['border'].get('left', {})),
            right=Side(**style_data['border'].get('right', {})),
            top=Side(**style_data['border'].get('top', {})),
            bottom=Side(**style_data['border'].get('bottom', {}))
        )
    if 'alignment' in style_data:
        cell.alignment = Alignment(**style_data['alignment'])
    if 'numFmt' in style_data:
        cell.number_format = style_data['numFmt']

def create_chart_from_spec(worksheet, chart_spec):
    """Crea un gráfico basado en la especificación del JSON."""
    chart_type = chart_spec.get('type', 'bar').lower()

    if chart_type in ['bar', 'col']:
        chart = BarChart()
        chart.type = chart_type
    elif chart_type == 'line':
        chart = LineChart()
    elif chart_type == 'pie':
        chart = PieChart()
    else:
        return

    chart.title = chart_spec.get('title', 'Gráfico sin Título')
    chart.style = chart_spec.get('style', 10)
    if 'y_axis_title' in chart_spec:
        chart.y_axis.title = chart_spec['y_axis_title']
    if 'x_axis_title' in chart_spec:
        chart.x_axis.title = chart_spec['x_axis_title']

    data = Reference(worksheet, range_string=chart_spec['data_range'])
    cats = Reference(worksheet, range_string=chart_spec['category_range'])

    chart.add_data(data, titles_from_data=chart_spec.get('titles_from_data', True))
    chart.set_categories(cats)
    
    worksheet.add_chart(chart, chart_spec.get('position', 'E1'))


# --- Endpoint Principal ---

@app.route('/create-excel', methods=['POST'])
def create_excel():
    try:
        # 1. Obtener y validar el JSON de la petición
        json_data = request.get_json()
        if not json_data or 'analysisData' not in json_data:
            return jsonify({"error": "El JSON no es válido o no contiene 'analysisData'."}), 400

        analysis_data = json_data.get('analysisData', [])
        conditional_rules = json_data.get('conditionalFormattingRules', [])
        chart_specs = json_data.get('charts', [])
        merge_cells_list = json_data.get('mergeCells', [])

        # 2. Crear el libro y la hoja de trabajo
        wb = Workbook()
        ws = wb.active
        ws.title = "Reporte Generado"

        # 3. Poblar celdas y aplicar estilos
        for row_details in analysis_data:
            for cell_data in row_details:
                if cell_data and 'address' in cell_data:
                    cell = ws[cell_data['address']]
                    cell.value = cell_data.get('value')
                    apply_styles_to_cell(cell, cell_data.get('style'))

        # 4. Aplicar combinación de celdas
        for cell_range in merge_cells_list:
            ws.merge_cells(cell_range)

        # 5. Aplicar formato condicional
        for rule_info in conditional_rules:
            style = rule_info.get('style', {})
            dxf = DifferentialStyle(
                font=Font(**style.get('font', {})),
                fill=PatternFill(**style.get('fill', {}))
            )
            
            rule_params = {'type': rule_info['type'], 'dxf': dxf}
            if 'operator' in rule_info:
                rule_params['operator'] = rule_info['operator']
            if 'formulae' in rule_info:
                rule_params['formula'] = rule_info['formulae']

            if rule_info['type'] == 'dataBar':
                rule = DataBarRule(
                    start_type='min', end_type='max', 
                    color=rule_info.get("color", "638EC6")
                )
            else:
                rule = Rule(**rule_params)

            ws.conditional_formatting.add(rule_info['ref'], rule)

        # 6. Crear e insertar los gráficos
        for spec in chart_specs:
            create_chart_from_spec(ws, spec)

        # 7. Ajustar ancho de columnas
        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) if max_length < 50 else 50
            ws.column_dimensions[column_letter].width = adjusted_width

        # 8. Guardar el archivo en memoria y enviarlo como respuesta
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name='reporte_completo.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        print(f"Error en /create-excel: {e}")
        return jsonify({"error": f"Error interno del servidor: {str(e)}"}), 500

# --- Punto de Entrada para Ejecutar la Aplicación ---
if __name__ == '__main__':
    app.run(debug=True, port=5000)
