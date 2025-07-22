from datetime import datetime, time
from openpyxl import load_workbook


class procesador_excel:
    def __init__(self):
        self.color_index_to_rgb = {
            3: "FF0000",  # Renuncia
            33: "00B0F0",  # Vacaciones
            47: "7030A0",  # Incapacidad
            43: "92D050",  # Renuncia
            6: "FFFF00",  # Despido
            49: "002060",  # Sancion o licencia no remunerada
            14: "00B050",  # Licencia maternidad
            16: "808080",  # No programado
            44: "FFC000",  # Licencia de luto
        }

        self.col_destino_color = {
            3: 104,
            33: 105,
            47: 106,
            43: 107,
            6: 108,
            49: 109,
            14: 110,
            16: 111,
            44: 112,
        }

    def contar_colores(self, ws, fila):
        colores = {idx: 0 for idx in self.color_index_to_rgb}
        for col in range(4, 66):  # Columnas de la 4 a la 65
            celda = ws.cell(row=fila, column=col)
            fill = celda.fill
            if fill and fill.start_color and fill.start_color.type == "rgb":
                color = fill.start_color.rgb
                if isinstance(color, str):
                    color = color[-6:].upper()
                    for idx, rgb in self.color_index_to_rgb.items():
                        if color == rgb:
                            colores[idx] += 1
        return colores

    def calcular_horas_y_colores(self, ruta_archivo):
        from openpyxl import load_workbook
        import os

        wb = load_workbook(ruta_archivo, keep_vba=True, data_only=True)
        ws = wb.worksheets[0]  # Siempre Hoja1

        valor_hora_extra = ws.cell(row=10, column=189).value or 0

        for fila in range(7, ws.max_row + 1):
            total_horas = 0
            total_extras = 0

            for idx, col in enumerate(range(4, 66, 4)):  # Columnas 4, 8, 12, ..., 64
                entrada = ws.cell(row=fila, column=col).value
                salida = ws.cell(row=fila, column=col + 1).value

                if isinstance(entrada, (datetime, time)) and isinstance(
                    salida, (datetime, time)
                ):
                    hora_entrada = entrada.hour + entrada.minute / 60
                    hora_salida = salida.hour + salida.minute / 60
                    horas_trabajadas = hora_salida - hora_entrada

                    if 6 <= hora_entrada < 12:
                        horas_trabajadas -= 1.5
                    elif 12 <= hora_entrada < 23:
                        horas_trabajadas -= 0.25

                    if horas_trabajadas < 0:
                        horas_trabajadas = 0

                    col_horas = 68 + idx  # Columnas 68 a 83
                    ws.cell(row=fila, column=col_horas).value = round(
                        horas_trabajadas / 24, 6
                    )
                    total_horas += horas_trabajadas / 24

                    # Extras mayores a 8h por día
                    if horas_trabajadas > 8:
                        total_extras += horas_trabajadas - 8

            # Total horas (col 84)
            ws.cell(row=fila, column=84).value = round(total_horas, 6)

            # Total extras (col 102)
            ws.cell(row=fila, column=102).value = round(total_extras, 2)

            # Calcular días trabajados (columna 103) contando celdas de horas > 0 entre columnas 68 y 83
            dias_trabajados = 0
            for c in range(68, 84):  # Columnas BP a CE
                valor = ws.cell(row=fila, column=c).value
                if isinstance(valor, (int, float)) and valor > 0:
                    dias_trabajados += 1
            ws.cell(row=fila, column=103).value = dias_trabajados

            # Sumar columnas 103 + 104 en la columna 113
            ws.cell(row=fila, column=113).value = f"=CY{fila}+CZ{fila}"

            # PAGO NORMAL o PARCIAL como valor numérico (columna 114)
            if dias_trabajados + total_extras > 15:
                ws.cell(row=fila, column=114).value = 711750
            else:
                ws.cell(row=fila, column=114).value = 711750

            # Valor horas extras (col 116)
            valor_total_extras = total_extras * valor_hora_extra
            ws.cell(row=fila, column=116).value = round(valor_total_extras, 2)

            # === Dominicales/Festivos ===
            festivos = sum(
                1 for col in range(4, 66) if ws.cell(row=fila, column=col).value == "F"
            )
            dominicales = sum(
                1 for col in range(4, 66) if ws.cell(row=fila, column=col).value == "D"
            )
            ws.cell(row=fila, column=117).value = festivos
            ws.cell(row=fila, column=118).value = dominicales

            # === Contar colores y escribir ===
            contadores = self.contar_colores(ws, fila)
            for idx, cantidad in contadores.items():
                col = self.col_destino_color[idx]
                ws.cell(row=fila, column=col).value = (
                    cantidad if cantidad != 0 else None
                )

            # === Ingresos y deducciones ===
            # ingresos = (valor_total_extras or 0) + (festivos + dominicales) * valor_hora_extra
            # deducciones = 0
            # if contadores[3] > 0 or contadores[6] > 0:  # Renuncia o despido
            #     deducciones += 50000
            # ws.cell(row=fila, column=120).value = round(ingresos, 2)
            # ws.cell(row=fila, column=121).value = round(deducciones, 2)

            # Total neto (col 124)
            # ws.cell(row=fila, column=124).value = round(ingresos - deducciones, 2)

            # Sumatoria ingresos (col 126) y total devengado (col 132)
            # ws.cell(row=fila, column=126).value = ingresos
            # ws.cell(row=fila, column=132).value = ingresos
            ws.cell(row=fila, column=122).value = (
                f"=DJ{fila}+DL{fila}+DN{fila}+DO{fila}"
            )
        # Guardar nuevo archivo
        base, ext = os.path.splitext(ruta_archivo)
        nuevo_nombre = f"{base}_Procesado{ext}"
        wb.save(nuevo_nombre)
        print(f"Archivo guardado como: {nuevo_nombre}")
