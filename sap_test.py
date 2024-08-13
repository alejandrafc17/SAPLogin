import win32com.client
import sys
import subprocess
import time

class SapGui:
    def __init__(self):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        self.start_sap_gui()
        time.sleep(5)  # Espera para que SAP GUI se inicie

        try:
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not isinstance(self.SapGuiAuto, win32com.client.CDispatch):
                raise Exception("SAP GUI no está disponible")

            application = self.SapGuiAuto.GetScriptingEngine
            self.connection = application.OpenConnection("PRODUCTIVO (PRD)", True)
            time.sleep(5)  # Espera para que la conexión se establezca

            self.session = self.connection.Children(0)
            self.session.findById("wnd[0]").maximize()

            # Realiza el login automáticamente
            self.sap_login()

            # Ejecuta la transacción ME5A
            self.execute_transaction("ME5A")

            # Realiza tareas en la transacción ME5A
            self.perform_me5a_tasks()

        except Exception as e:
            print(f"Error en la inicialización de SAP GUI: {e}")

    def start_sap_gui(self):
        """Inicia SAP GUI"""
        try:
            subprocess.Popen(self.path)
        except Exception as e:
            print(f"Error al iniciar SAP GUI: {e}")
            sys.exit(1)

    def sap_login(self):
        """Realiza el login en SAP"""
        try:
            self.wait_for_element("wnd[0]/usr/txtRSYST-MANDT", timeout=5)

            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "300"
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "SUP10149"
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "PLANEACION2024*08"
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "ES"

            self.session.findById("wnd[0]").sendVKey(0)

            print("Login exitoso")
        except Exception as e:
            print(f"Error en el login de SAP: {e}")

    def execute_transaction(self, transaction_code):
        """Ejecuta una transacción en SAP"""
        try:
            self.wait_for_element("wnd[0]/tbar[0]/okcd", timeout=5)
            self.session.findById("wnd[0]/tbar[0]/okcd").text = transaction_code
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(2)  # Breve pausa para permitir la carga de la transacción
            print(f"Transacción {transaction_code} ejecutada exitosamente")
        except Exception as e:
            print(f"Error al ejecutar la transacción {transaction_code}: {e}")

    def perform_me5a_tasks(self):
        """Realiza tareas específicas dentro de la transacción ME5A"""
        try:
            # Ingresar rango de materiales
            self.enter_data("wnd[0]/usr/ctxtBA_MATNR-LOW", "110000000")  # Material desde
            self.enter_data("wnd[0]/usr/ctxtBA_MATNR-HIGH", "129999999")  # Material hasta

            # Ingresar el Alcance de lista
            self.enter_data("wnd[0]/usr/ctxtP_LSTUB", "ALV")  # Alcance de lista
            
            # Marcar la casilla de verificación "P_FREIG"
            self.session.findById("wnd[0]/usr/chkP_ERLBA").setFocus()
            self.session.findById("wnd[0]/usr/chkP_ERLBA").selected = False
                
            # Marcar la casilla de verificación "P_FREIG"
            self.session.findById("wnd[0]/usr/chkP_FREIG").setFocus()
            self.session.findById("wnd[0]/usr/chkP_FREIG").selected = False
            
            # Presionar el botón "Ejecutar"
            self.click_button("wnd[0]/tbar[1]/btn[8]")  # Botón de ejecutar

            print("Tareas dentro de ME5A completadas")
        except Exception as e:
            print(f"Error al realizar tareas en ME5A: {e}")

    def enter_data(self, element_id, value):
        """Ingresa datos en un campo específico"""
        try:
            self.wait_for_element(element_id, timeout=5)
            self.session.findById(element_id).text = value
        except Exception as e:
            print(f"Error al ingresar datos en {element_id}: {e}")

    def click_button(self, button_id):
        """Hace clic en un botón específico"""
        try:
            self.wait_for_element(button_id, timeout=5)
            self.session.findById(button_id).press()
        except Exception as e:
            print(f"Error al hacer clic en el botón {button_id}: {e}")

    def wait_for_element(self, element_id, timeout=10):
        """Espera a que un elemento esté disponible"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                if self.session.findById(element_id):
                    return
            except:
                pass
            time.sleep(1)
        raise TimeoutError(f"No se pudo encontrar el elemento {element_id} en el tiempo especificado.")

if __name__ == '__main__':
    SapGui()  # Crea una instancia de la clase y realiza el login automáticamente
