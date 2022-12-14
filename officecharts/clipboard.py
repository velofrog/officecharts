import io
import ctypes

os_clipboard = None


def _set_os_clipboard(shared_library: ctypes.CDLL) -> None:
    global os_clipboard
    os_clipboard = shared_library

def send_officedrawing(data: io.BytesIO) -> None:
    if os_clipboard:
        os_clipboard.send_officedrawing_to_clipboard(data.getbuffer().nbytes, data.getvalue())