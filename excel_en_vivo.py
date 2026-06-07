"""
EXCEL EN VIVO - Lee COM8 y muestra 12 temperaturas en Excel
Version robusta con mutex, auto-save, reconexion y reanudacion
"""
import serial, time, os, sys
import win32com.client

COM = "COM8"
BAUD = 9600
MAX_TEMP = 200
MIN_TEMP = -50
DIR = r"C:\Users\Acer\Desktop\Arduino_Excel_Logger"
STATUS = os.path.join(DIR, "status.txt")
SAVE_PATH = os.path.join(DIR, "datos_temperatura.xlsx")
AUTOSAVE_INTERVAL = 50

HEADERS = ["Timestamp","Local_1","Local_2","Local_3","Local_4","Local_5","Local_6",
           "Remoto_1","Remoto_2","Remoto_3","Remoto_4","Remoto_5","Remoto_6"]

excel = None
wb = None

def filtrar(v):
    try:
        vf = float(v.strip().lower().replace("nan", "0"))
        return round(vf, 2) if MIN_TEMP <= vf <= MAX_TEMP else 0.0
    except:
        return 0.0

def single_instance():
    try:
        import win32event, win32api, winerror
        h = win32event.CreateMutex(None, False, "Arduino_Excel_Logger_Mutex")
        err = win32api.GetLastError()
        if err == winerror.ERROR_ALREADY_EXISTS:
            r = win32event.WaitForSingleObject(h, 0)
            if r == winerror.WAIT_TIMEOUT:
                print("ERROR: Otra instancia esta corriendo.")
                sys.exit(1)
    except:
        pass

def conectar(retries=10, backoff=2):
    last_err = None
    for att in range(1, retries + 1):
        try:
            ser = serial.Serial(COM, BAUD, timeout=0.5)
            for _ in range(5):
                ser.setDTR(False)
                time.sleep(0.3)
                ser.setDTR(True)
                time.sleep(0.8)
            ser.setDTR(False)
            time.sleep(0.3)
            ser.setDTR(True)
            time.sleep(5)
            print(f"  COM8 conectado (intento {att})")
            return ser
        except serial.SerialException as e:
            last_err = e
            print(f"  COM8 fallo (intento {att}/{retries}): {e}")
            if att < retries:
                time.sleep(backoff * att)
    raise RuntimeError(f"No se pudo conectar a {COM} tras {retries} intentos: {last_err}")

def abrir_excel():
    global excel, wb
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.WindowState = -4137
    if os.path.exists(SAVE_PATH):
        try:
            wb = excel.Workbooks.Open(SAVE_PATH)
            ws = wb.ActiveSheet
            last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            fila = last + 1 if last >= 1 else 2
            print(f"  Reanudado desde {SAVE_PATH}, fila {fila}")
            return excel, wb, ws, fila
        except:
            pass
    wb = excel.Workbooks.Add()
    ws = wb.ActiveSheet
    ws.Name = "Temperaturas"
    formatear(ws, excel)
    print("  Nuevo workbook creado")
    return excel, wb, ws, 2

def formatear(ws, excel):
    xlCenter = -4108
    for i, h in enumerate(HEADERS, 1):
        c = ws.Cells(1, i)
        c.Value = h
        c.Font.Bold = True
        c.Font.Color = 0xFFFFFF
        c.ColumnWidth = 14
        c.HorizontalAlignment = xlCenter
    ws.Range("A1").Interior.Color = 0x4682B4
    ws.Range("B1:G1").Interior.Color = 0x3CB371
    ws.Range("H1:M1").Interior.Color = 0xFFA500
    ws.Range("A2").Select()
    try:
        excel.ActiveWindow.FreezePanes = True
    except:
        pass

def guardar():
    try:
        if os.path.exists(SAVE_PATH):
            wb.Save()
        else:
            wb.SaveAs(SAVE_PATH)
    except:
        pass

def cerrar_excel():
    global excel, wb
    try:
        guardar()
    except:
        pass
    try:
        wb.Close(SaveChanges=False)
    except:
        pass
    try:
        excel.Quit()
    except:
        pass

def main():
    single_instance()

    with open(STATUS, "w") as sf:
        sf.write("INICIANDO\n")

    print("Conectando a COM8...")
    ser = conectar()

    print("Abriendo Excel...")
    excel, wb, ws, fila = abrir_excel()

    contador = 0
    print("Capturando datos en Excel... Ctrl+C para detener")

    try:
        while True:
            try:
                linea = ser.readline()
            except:
                print("  Serial perdido, reconectando...")
                try:
                    ser.close()
                except:
                    pass
                ser = conectar(retries=5, backoff=3)
                continue
            if not linea:
                continue

            try:
                texto = linea.decode("utf-8").strip()
            except:
                continue
            if not texto:
                continue

            partes = [p.strip() for p in texto.split(",")]
            if len(partes) < 12:
                continue
            try:
                float(partes[0])
            except:
                continue

            ts = time.strftime("%Y-%m-%d %H:%M:%S")
            try:
                ws.Cells(fila, 1).Value = ts
                for i in range(12):
                    ws.Cells(fila, i + 2).Value = filtrar(partes[i])
            except:
                try:
                    print("  Excel perdido, reconectando...")
                    excel, wb, ws, fila = abrir_excel()
                    ws.Cells(fila, 1).Value = ts
                    for i in range(12):
                        ws.Cells(fila, i + 2).Value = filtrar(partes[i])
                except:
                    pass

            contador += 1
            if contador % 10 == 0:
                print(f"  {contador} registros")
                with open(STATUS, "w") as sf:
                    sf.write(f"{contador} registros | Fila {fila} | Ultimo: {ts}\n")
            if contador % AUTOSAVE_INTERVAL == 0:
                guardar()
                print(f"  Auto-guardado en {SAVE_PATH}")

            fila += 1

    except KeyboardInterrupt:
        print("\nDetenido por usuario")
    except Exception as e:
        print(f"\nError critico: {e}")
    finally:
        try:
            ser.close()
        except:
            pass
        cerrar_excel()
        with open(STATUS, "w") as sf:
            sf.write(f"DETENIDO - Total: {contador} registros\n")
        print(f"Total: {contador} registros")

if __name__ == "__main__":
    main()
