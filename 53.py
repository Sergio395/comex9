import os
import pandas as pd
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from babel.numbers import format_currency
import shutil #207

class ComexOperation:
    def __init__(self, fecha_ingreso, referencia, otorgante, monto_total, proporcion_del_total_por_fc, moneda, importe_factura,
                 porcentaje_factura, numero_factura, tipo_cambio_factura, tc_acreditacion, fecha_liquidacion,
                 doc_ajuste, rc, op1, op2, obs, ret_ext):
        self.fecha_ingreso = fecha_ingreso
        self.referencia = int(referencia) if referencia else None
        self.otorgante = otorgante
        self.monto_total = float(monto_total) if monto_total else None
        self.proporcion_del_total_por_fc = float(proporcion_del_total_por_fc) if proporcion_del_total_por_fc else None
        self.moneda = moneda
        self.importe_factura = float(importe_factura) if importe_factura else None
        self.porcentaje_factura = float(porcentaje_factura) if porcentaje_factura else None
        self.numero_factura = numero_factura
        self.tipo_cambio_factura = float(tipo_cambio_factura) if tipo_cambio_factura else None
        self.tc_acreditacion = float(tc_acreditacion) if tc_acreditacion else None
        self.fecha_liquidacion = fecha_liquidacion  # Nuevo campo
        self.ret_ext = float(ret_ext) if ret_ext else 0.0
        self.monto_acreditado = self.calcular_monto_acreditado() if self.importe_factura and self.tc_acreditacion else None
        self.gasto_usd = self.calcular_gasto_usd()
        self.ajuste = self.calcular_ajuste()
        self.doc_ajuste = doc_ajuste
        self.rc = rc
        self.op1 = op1
        self.op2 = op2
        self.obs = obs

    def calcular_monto_acreditado(self):
        if self.proporcion_del_total_por_fc == 1.0:
            monto_acreditado = self.monto_total * self.tc_acreditacion
        elif self.proporcion_del_total_por_fc == 0.8:
            monto_acreditado = (self.monto_total * 0.8) * self.tc_acreditacion
        elif self.proporcion_del_total_por_fc == 0.2:
            monto_acreditado = (self.monto_total * 0.2) * self.tc_acreditacion
        return float(monto_acreditado - self.ret_ext) if self.ret_ext else float(monto_acreditado)

    def calcular_gasto_usd(self):
        if self.importe_factura is not None and self.porcentaje_factura is not None:
            importe_factura_porcentaje = self.importe_factura * self.porcentaje_factura
            return float(importe_factura_porcentaje * self.tc_acreditacion) - self.monto_acreditado if self.monto_acreditado is not None else 0.0
        return 0.0

    def calcular_ajuste(self):
        if self.importe_factura is not None and self.porcentaje_factura is not None:
            importe_factura_porcentaje = self.importe_factura * self.porcentaje_factura
            return float(self.monto_acreditado - ((importe_factura_porcentaje * self.tipo_cambio_factura) - self.gasto_usd)) if self.monto_acreditado is not None else 0.0
        return 0.0

class ComexApp:
    def __init__(self):
        self.data = []
        self.file_path_all_fields = 'Liquidacion_Cobranza_Exterior_nueva.xlsx'  # Nombre del archivo Excel
        self.load_data()

    def load_data(self):
        if os.path.exists(self.file_path_all_fields):
            df = pd.read_excel(self.file_path_all_fields)
            for _, row in df.iterrows():
                operacion = ComexOperation(
                    row['Fecha ingreso'], row['Referencia'], row['Otorgante'], row['Monto Total'], 
                    row['Proporcion del total por fc'], row['Moneda'], row['Importe Factura'], row['Porcentaje Factura'], 
                    row['Número Factura'], row['Tipo de Cambio Factura'], row['TC Acreditacion'],
                    row.get('Fecha liquidacion', None),  # Nueva columna
                    row['Doc Ajuste'], row['RC'], row['Op1'], row['Op2'], row['Obs'], row['Ret Ext']
                )
                self.data.append(operacion)

    def guardar_operacion(self):
        ref = input("Referencia (entero): ")
        if any(op.referencia == int(ref) for op in self.data):
            print("Advertencia: La referencia ya existe.")
            return

        operacion = ComexOperation(
            input("Fecha ingreso: "), ref, input("Otorgante: "), input("Monto Total: "),
            input("Proporcion del total por fc (1, 0.8, 0.2): "), input("Moneda: "), input("Importe Factura: "),
            input("Porcentaje Factura (1, 0.8, 0.2): "), input("Número Factura: "),
            input("Tipo de Cambio Factura: "), input("TC Acreditacion: "), input("Fecha liquidacion: "),
            input("Doc Ajuste: "), input("RC: "), input("Op1: "), input("Op2: "), input("Obs: "), input("Ret Ext: ")
        )
        self.data.append(operacion)
        self.export_to_excel()

        # Mostrar campos calculados
        print(f"Monto Acreditado: {operacion.monto_acreditado}")
        print(f"Gasto USD: {operacion.gasto_usd}")
        print(f"Ajuste: {operacion.ajuste}")

    def modificar_operacion(self):
        ref = input("Referencia a modificar (entero): ")
        operacion = next((op for op in self.data if op.referencia == int(ref)), None)
        if not operacion:
            print("Advertencia: La referencia no existe.")
            return
        
        operacion.fecha_ingreso = input(f"Fecha ingreso ({operacion.fecha_ingreso}): ") or operacion.fecha_ingreso
        operacion.otorgante = input(f"Otorgante ({operacion.otorgante}): ") or operacion.otorgante
        operacion.monto_total = float(input(f"Monto Total ({operacion.monto_total}): ") or operacion.monto_total)
        operacion.proporcion_del_total_por_fc = float(input(f"Proporcion del total por fc ({operacion.proporcion_del_total_por_fc}): ") or operacion.proporcion_del_total_por_fc)
        operacion.moneda = input(f"Moneda ({operacion.moneda}): ") or operacion.moneda
        operacion.importe_factura = float(input(f"Importe Factura ({operacion.importe_factura}): ") or operacion.importe_factura)
        operacion.porcentaje_factura = float(input(f"Porcentaje Factura ({operacion.porcentaje_factura}): ") or operacion.porcentaje_factura)
        operacion.numero_factura = input(f"Número Factura ({operacion.numero_factura}): ") or operacion.numero_factura
        operacion.tipo_cambio_factura = float(input(f"Tipo de Cambio Factura ({operacion.tipo_cambio_factura}): ") or operacion.tipo_cambio_factura)
        
        tc_acreditacion_input = input(f"TC Acreditacion ({operacion.tc_acreditacion}): ")
        operacion.tc_acreditacion = float(tc_acreditacion_input) if tc_acreditacion_input else operacion.tc_acreditacion

        operacion.fecha_liquidacion = input(f"Fecha liquidacion ({operacion.fecha_liquidacion}): ") or operacion.fecha_liquidacion  # Nuevo campo
        
        operacion.doc_ajuste = input(f"Doc Ajuste ({operacion.doc_ajuste}): ") or operacion.doc_ajuste
        operacion.rc = input(f"RC ({operacion.rc}): ") or operacion.rc
        operacion.op1 = input(f"Op1 ({operacion.op1}): ") or operacion.op1
        operacion.op2 = input(f"Op2 ({operacion.op2}): ") or operacion.op2
        operacion.obs = input(f"Obs ({operacion.obs}): ") or operacion.obs
        operacion.ret_ext = float(input(f"Ret Ext ({operacion.ret_ext}): ") or operacion.ret_ext)

        operacion.monto_acreditado = operacion.calcular_monto_acreditado()
        operacion.gasto_usd = operacion.calcular_gasto_usd()
        operacion.ajuste = operacion.calcular_ajuste()

        self.export_to_excel()

        # Mostrar campos calculados
        print(f"Monto Acreditado: {operacion.monto_acreditado}")
        print(f"Gasto USD: {operacion.gasto_usd}")
        print(f"Ajuste: {operacion.ajuste}")

    def listar_operaciones(self):
        for operacion in self.data:
            print(f"Fecha ingreso: {operacion.fecha_ingreso}")
            print(f"Referencia: {operacion.referencia}")
            print(f"Otorgante: {operacion.otorgante}")
            print(f"Monto Total: {operacion.monto_total}")
            print(f"Proporcion del total por fc: {operacion.proporcion_del_total_por_fc}")
            print(f"Moneda: {operacion.moneda}")
            print(f"Importe Factura: {operacion.importe_factura}")
            print(f"Porcentaje Factura: {operacion.porcentaje_factura}")
            print(f"Número Factura: {operacion.numero_factura}")
            print(f"Tipo de Cambio Factura: {operacion.tipo_cambio_factura}")
            print(f"TC Acreditacion: {operacion.tc_acreditacion}")
            print(f"Fecha liquidacion: {operacion.fecha_liquidacion}")  # Nuevo campo
            print(f"Doc Ajuste: {operacion.doc_ajuste}")
            print(f"RC: {operacion.rc}")
            print(f"Op1: {operacion.op1}")
            print(f"Op2: {operacion.op2}")
            print(f"Obs: {operacion.obs}")
            print(f"Ret Ext: {operacion.ret_ext}")
            print(f"Monto Acreditado: {operacion.monto_acreditado}")
            print(f"Gasto USD: {operacion.gasto_usd}")
            print(f"Ajuste: {operacion.ajuste}")
            print("-------------------------")

    def buscar_por_referencia(self):
        ref = input("Referencia a buscar (entero): ")
        ref = int(ref)
        if os.path.exists(self.file_path_all_fields):
            df = pd.read_excel(self.file_path_all_fields)
            if 'Referencia' in df.columns:
                matching_rows = df[df['Referencia'] == ref]
                if not matching_rows.empty:
                    row = matching_rows.iloc[0]
                    for column in row.index:
                        print(f"{column}: {row[column] if pd.notnull(row[column]) else ''}")
                else:
                    print("No se encontró ninguna operación con la referencia proporcionada.")
            else:
                print("La columna 'Referencia' no se encuentra en el archivo Excel.")
        else:
            print("No se encontró el archivo Excel con los datos.")

    def export_to_excel(self):
        all_fields_data = {
            "Fecha ingreso": [op.fecha_ingreso for op in self.data],
            "Referencia": [op.referencia for op in self.data],
            "Otorgante": [op.otorgante for op in self.data],
            "Monto Total": [op.monto_total for op in self.data],
            "Proporcion del total por fc": [op.proporcion_del_total_por_fc for op in self.data],
            "Moneda": [op.moneda for op in self.data],
            "Importe Factura": [op.importe_factura for op in self.data],
            "Porcentaje Factura": [op.porcentaje_factura for op in self.data],
            "Número Factura": [op.numero_factura for op in self.data],
            "Tipo de Cambio Factura": [op.tipo_cambio_factura for op in self.data],
            "TC Acreditacion": [op.tc_acreditacion for op in self.data],
            "Fecha liquidacion": [op.fecha_liquidacion for op in self.data],  # Nueva columna
            "Monto Acreditado": [op.monto_acreditado for op in self.data],
            "Gasto USD": [op.gasto_usd for op in self.data],
            "Ajuste": [op.ajuste for op in self.data],
            "Doc Ajuste": [op.doc_ajuste for op in self.data],
            "RC": [op.rc for op in self.data],
            "Op1": [op.op1 for op in self.data],
            "Op2": [op.op2 for op in self.data],
            "Obs": [op.obs for op in self.data],
            "Ret Ext": [op.ret_ext for op in self.data],
        }

        df_all_fields = pd.DataFrame(all_fields_data)
        df_all_fields.to_excel(self.file_path_all_fields, index=False)
        ruta_origen = 'G:/Mi unidad/001ComercioExterior/'
        ruta_destino = 'G:/Unidades compartidas/Tesoreria/COMERCIO EXTERIOR/'
        shutil.copy(ruta_origen + 'Liquidacion_Cobranza_Exterior_nueva.xlsx', ruta_destino)
    def generar_planilla_cobro(self):
        ref = input("Ingrese la referencia para generar la planilla de cobro (entero): ")
        ref = int(ref)
        if os.path.exists(self.file_path_all_fields):
            df = pd.read_excel(self.file_path_all_fields)
            if 'Referencia' in df.columns:
                matching_rows = df[df['Referencia'] == ref]
                if not matching_rows.empty:
                    row = matching_rows.iloc[0]
                
                    data_to_export = {
                        "Fecha ingreso": [row['Fecha ingreso']],
                        "Importe total en banco": [row['Monto Total']*row['Proporcion del total por fc']*row["TC Acreditacion"]],
                        "Proporcion del total por fc": [row['Proporcion del total por fc']],
                        #"Gs bancarios USD": [row['Gasto USD']/row["TC Acreditacion"]],
                        "Gs bancarios USD": 0,
                        "Otorgante": [row['Otorgante']],
                        "TC Acreditacion": [row['TC Acreditacion']],
                        "Importe en USD": ([row['Monto Total']*row['Proporcion del total por fc']*row["TC Acreditacion"]]/row["TC Acreditacion"]),
                        "Tipo op:Pesif al _ %": [(row ['Proporcion del total por fc']*100)] # Ser
                    }

                    df_planilla_cobro = pd.DataFrame(data_to_export)
                    file_path_planilla_cobro = f'Planilla_cobro_ME.xlsx'

                    df_planilla_cobro.to_excel(file_path_planilla_cobro, index=False)
                    print(f"Se ha generado exitosamente la planilla de cobro en '{file_path_planilla_cobro}'.")
                else:
                    print("No se encontró ninguna operación con la referencia proporcionada.")
            else:
                print("La columna 'Referencia' no se encuentra en el archivo Excel.")
        else:
            print("No se encontró el archivo Excel con los datos.")

def menu():
    app = ComexApp()
    while True:
        print("1. Alta")
        print("2. Modificación")
        print("3. Buscar por Referencia")
        print("4. Listar")
        print("5. Generar Planilla para cobro en ME")
        print("6. Salir")
        choice = input("Seleccione una opción: ")
        if choice == '1':
            app.guardar_operacion()
        elif choice == '2':
            app.modificar_operacion()
        elif choice == '3':
            app.buscar_por_referencia()
        elif choice == '4':
            app.listar_operaciones()
        elif choice == '5':
            app.generar_planilla_cobro()
        elif choice == '6':
            break
        else:
            print("Opción no válida, por favor intente de nuevo.")

if __name__ == "__main__":
    menu()
