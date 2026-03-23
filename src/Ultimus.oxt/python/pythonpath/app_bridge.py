import ctypes
import time
import sys

# --- Funzioni principali ---

def trova_autocad(nome_finestra):
    """
    Cerca una finestra di AutoCAD che contenga 'nome_finestra' nel titolo.
    
    Argomenti:
        nome_finestra (str): parte del titolo della finestra da cercare.
    
    Ritorna:
        hwnd (int) → handle della finestra trovata, oppure None.
    """
    user32 = ctypes.WinDLL('user32')
    buffer = ctypes.create_unicode_buffer(255)
    hwnd = ctypes.c_void_p()

    def enum_windows_proc(h, lparam):
        user32.GetWindowTextW(h, buffer, 255)
        titolo = buffer.value
        if nome_finestra.lower() in titolo.lower():
            hwnd.value = h
            return False  # interrompe la ricerca
        return True  # continua

    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    user32.EnumWindows(EnumWindowsProc(enum_windows_proc), 0)
    return hwnd.value if hwnd.value else None


def attiva_finestra(hwnd):
    if hwnd:
        user32 = ctypes.WinDLL('user32')
        SW_RESTORE = 9
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)
        time.sleep(0.3)
        return True
    return False


def invia_tasto(ch):
    user32 = ctypes.WinDLL('user32')
    vk = user32.VkKeyScanW(ord(ch))
    user32.keybd_event(vk, 0, 0, 0)
    user32.keybd_event(vk, 0, 2, 0)
    time.sleep(0.05)


def invia_testo(testo):
    for ch in testo:
        invia_tasto(ch)
    user32 = ctypes.WinDLL('user32')
    user32.keybd_event(0x0D, 0, 0, 0)
    user32.keybd_event(0x0D, 0, 2, 0)


def autocad(nome_finestra, comando):
    """
    Esegue un comando su una finestra AutoCAD specifica.

    Argomenti:
        nome_finestra (str): parte del titolo della finestra.
        comando (str): comando AutoCAD da inviare.
    """
    hwnd = trova_autocad(nome_finestra)
    if attiva_finestra(hwnd):
        invia_testo(comando)
        print(f"✅ Comando '{comando}' inviato alla finestra '{nome_finestra}'.")
    else:
        print(f"❌ Nessuna finestra AutoCAD trovata con '{nome_finestra}' nel titolo.")


# --- Esecuzione da riga di comando ---
# if __name__ == "__main__":
#     if len(sys.argv) < 3:
#         print("Uso: python script.py <nome_finestra> <comando>")
#         sys.exit(1)
#     nome_finestra = sys.argv[1]
#     comando = sys.argv[2]
#     autocad(nome_finestra, comando)
