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

HWND create_window() {
    return CreateWindowExA(0, "Static", NULL, 0,
                           0, 0, 0, 0,
                           NULL, NULL, NULL, NULL);
}

PyObject *send_officedrawing_to_clipboard(PyObject *self, PyObject *args) {
    PyObject *bytes_object;

    if (!PyArg_ParseTuple(args, "S", &bytes_object)) {
        return NULL;
    }

    std::size_t length = PyBytes_Size(bytes_object);
    char *data = PyBytes_AsString(bytes_object);

    HGLOBAL hglbData;
    if (length == 0) {
        std::cerr << "Nothing to send to clipboard\n";
        return Py_BuildValue("");
    }

    hglbData = GlobalAlloc(GMEM_MOVEABLE|GMEM_SHARE|GMEM_ZEROINIT, length);
    if (hglbData == NULL) {
        std::cerr << "Error sending data to clipboard: Failed to allocate memory\n";
        return Py_BuildValue("");
    }

    uint8_t *ptr = (uint8_t *)GlobalLock(hglbData);
    if (ptr == NULL) {
        std::cerr << "Error sending data to clipboard: Failed to get pointer to memory object\n";
        return Py_BuildValue("");
    }

    std::memcpy(ptr, data, length);

    GlobalUnlock(hglbData);

    // Not creating our own window sometimes causes setClipboardData to silently fail
    //    i.e. appears to succeed, but no actual data placed on clipboard
    HWND clipboardWnd = create_window();
    if (clipboardWnd == NULL) {
        std::cerr << "Could not create temporary clipboard window\n";
        return Py_BuildValue("");
    }

    if (OpenClipboard(clipboardWnd)) {
        if (EmptyClipboard()) {
            if (SetClipboardData(officedrawing_clipboardformat(), hglbData)) {
                std::cout << "Sent " << length << " bytes to clipboard\n";
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
    DestroyWindow(clipboardWnd);

    return Py_BuildValue("");
}

static PyObject *test(PyObject *self, PyObject *args) {
    std::cout << "hello world.should return PyBuild('')\n";
    return Py_BuildValue("");
}

static PyMethodDef os_clipboard_methods[] =
{
    {"test", test, METH_VARARGS, "Test function"},
    {"send_officedrawing_to_clipboard", send_officedrawing_to_clipboard, METH_VARARGS,
        "Send OfficeDrawing to Windows Clipboard"},
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
