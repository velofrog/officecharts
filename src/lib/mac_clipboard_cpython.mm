#define PY_SSIZE_T_CLEAN
#include <Python.h>
#import <Foundation/Foundation.h>
#import <AppKit/NSPasteboard.h>
#include <vector>
#include <string>
#include <iostream>
#include <cstddef>


extern "C" {
    __attribute__((visibility("default"))) PyMODINIT_FUNC PyInit_os_clipboard(void);
}

bool officedrawing_uti(std::string &uti) {
    CFStringRef dyn_uti;
    CFStringRef tag = CFSTR("com.microsoft.Art--GVML-ClipFormat");
    bool result = false;

    dyn_uti = UTTypeCreatePreferredIdentifierForTag(kUTTagClassNSPboardType, tag, kUTTypeContent);
    if (dyn_uti) {
        const char *c_str = CFStringGetCStringPtr(dyn_uti, kCFStringEncodingUTF8);
        if (c_str) {
            uti.assign(c_str);
            result = true;
        }
        CFRelease(dyn_uti);
    }

    return result;
}

PyObject *send_officedrawing_to_clipboard(PyObject *self, PyObject *args) {
    PyObject *bytes_object;

    if (!PyArg_ParseTuple(args, "S", &bytes_object)) {
        return NULL;
    }

    std::size_t length = PyBytes_Size(bytes_object);
    char *data = PyBytes_AsString(bytes_object);
    if ((length == 0) || (data == NULL)) {
        std::cerr << "Nothing to send to clipboard\n";
        return Py_BuildValue("");
    }

    std::string uti;
    if (!officedrawing_uti(uti)) {
        std::cerr << "Failed to send to clipboard. Could not create preferred identifier tag.\n";
        return Py_BuildValue("");
    }

    @autoreleasepool {
        NSPasteboard *pbpaste = [NSPasteboard generalPasteboard];
        NSString *ns_uti = [NSString stringWithUTF8String:uti.c_str()];
        [pbpaste declareTypes: [NSArray arrayWithObject:ns_uti] owner:nil];

        NSData *pbdata = [NSData dataWithBytesNoCopy:data length:length freeWhenDone:false];

        [pbpaste clearContents];
        [pbpaste setData:pbdata forType:ns_uti];
        std::cout << "Uploaded " << length << " bytes to clipboard\n";
    }

    return Py_BuildValue("");
}

static PyMethodDef os_clipboard_methods[] =
{
    {"send_officedrawing_to_clipboard", send_officedrawing_to_clipboard, METH_VARARGS,
     "Places an OfficeDrawing object on the Mac Clipboard"},
    { NULL, NULL, 0, NULL}
};

static struct PyModuleDef os_clipboardmodule = {
    PyModuleDef_HEAD_INIT,
    "os_clipboard",
    "Places an OfficeDrawing object on the Mac clipboard",
    -1,
    os_clipboard_methods
};

PyMODINIT_FUNC PyInit_os_clipboard(void) {
    return PyModule_Create(&os_clipboardmodule);
}
