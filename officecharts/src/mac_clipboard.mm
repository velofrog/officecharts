#import <Foundation/Foundation.h>
#import <AppKit/NSPasteboard.h>
#include <vector>
#include <string>
#include <iostream>
#include <cstddef>


extern "C" __attribute__((visibility("default"))) void send_officedrawing_to_clipboard(
    std::size_t length, std::uint8_t *data);

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
