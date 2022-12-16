#define PY_SSIZE_T_CLEAN
#include <Python.h>
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <vector>
#include <iostream>
#include <cstddef>

extern "C" {
    static PyObject *test(PyObject *self, PyObject *args);
    PyMODINIT_FUNC PyInit_os_clipboard(void);
}

uint32_t CF_OFFICEDRAWING_OBJECT = 0;

uint32_t officedrawing_clipboardformat() {
    if (CF_OFFICEDRAWING_OBJECT == 0) {
        CF_OFFICEDRAWING_OBJECT = RegisterClipboardFormatA("Art::GVML ClipFormat");
    }
    return CF_OFFICEDRAWING_OBJECT;
}

void send_officedrawing_to_clipboard(std::size_t length, std::uint8_t *data) {
    HGLOBAL hglbData;
    if (length == 0) return;

    hglbData = GlobalAlloc(GMEM_MOVEABLE|GMEM_ZEROINIT, length);
    if (hglbData == NULL) {
        std::cerr << "Error sending data to clipboard: Failed to allocate memory\n";
        return;
    }

    uint8_t *ptr = (uint8_t *)GlobalLock(hglbData);
    if (ptr == NULL) {
        std::cerr << "Error sending data to clipboard: Failed to get pointer to memory object\n";
        return;
    }

    std::memcpy(ptr, data, length);

    GlobalUnlock(hglbData);

    if (OpenClipboard(NULL)) {
        if (EmptyClipboard()) {
            if (SetClipboardData(officedrawing_clipboardformat(), hglbData)) {
                std::cout << "Uploaded " << length << " bytes to clipboard\n";
            } else {
                std::cerr << "Error placing data on clipboard\n";
            }
        } else {
            std::cerr << "Error clearing clipboard\n";
        }
        CloseClipboard();
    } else {
        std::cerr << "Error opening clipboard\n";
    }

    GlobalFree(hglbData);
}

static PyObject *test(PyObject *self, PyObject *args) {
    std::cout << "hello world.should return PyBuild('')\n";
    return Py_BuildValue("");
}

static PyMethodDef os_clipboard_methods[] =
{
    {"test", test, METH_VARARGS, "Test function"},
    { NULL, NULL, 0, NULL}
};

static struct PyModuleDef os_clipboardmodule = {
    PyModuleDef_HEAD_INIT,
    "os_clipboard",
    "Sends an Office Drawing Chart to the Windows clipboard",
    -1,
    os_clipboard_methods
};

PyMODINIT_FUNC PyInit_os_clipboard(void) {
    return PyModule_Create(&os_clipboardmodule);
}
