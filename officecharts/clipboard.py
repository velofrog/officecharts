import io
import ctypes

os_clipboard = None


def _set_os_clipboard(shared_library: ctypes.CDLL) -> None:
    global os_clipboard
    print(f"setting os_clipboard to {shared_library}")
    os_clipboard = shared_library

def send_officedrawing(data: io.BytesIO) -> None:
    print(f"os_clipboard={os_clipboard}")
    if os_clipboard:
        os_clipboard.send_officedrawing_to_clipboard(data.getbuffer().nbytes, data.getvalue())