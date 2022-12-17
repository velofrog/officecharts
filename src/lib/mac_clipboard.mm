#include <Python.h>
#import <Foundation/Foundation.h>
#import <AppKit/NSPasteboard.h>
#include <vector>
#include <string>
#include <iostream>
#include <cstddef>


extern "C" {
    __attribute__((visibility("default"))) void send_officedrawing_to_clipboard(
        std::size_t length, std::uint8_t *data);
    __attribute__((visibility("default"))) static PyObject *test(PyObject *self, PyObject *args);
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

void send_officedrawing_to_clipboard(std::size_t length, std::uint8_t *data) {
    std::string uti;
    if (!officedrawing_uti(uti)) {
        std::cerr << "Failed to send to clipboard. Could not create preferred identifier tag.\n";
        return;
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
}

static PyObject *test(PyObject *self, PyObject *args) {
    std::cout << "hello world\n";
    return NULL;
}

static PyMethodDef os_clipboard_methods[] =
{
    {"test", test, METH_VARARGS, "Test function"},
    { NULL, NULL, 0, NULL}
};

static struct PyModuleDef os_clipboardmodule = {
    PyModuleDef_HEAD_INIT,
    "os_clipboard",
    "Sends an Office Drawing Chart to the Mac clipboard",
    -1,
    os_clipboard_methods
};

PyMODINIT_FUNC PyInit_os_clipboard(void) {
    return PyModule_Create(&os_clipboardmodule);
}
