#import <Foundation/Foundation.h>
#import <Vision/Vision.h>
#import <PDFKit/PDFKit.h>
#import <ImageIO/ImageIO.h>
#import <CoreGraphics/CoreGraphics.h>

static NSString *recognize(CGImageRef image, NSError **error) {
    VNImageRequestHandler *handler = [[VNImageRequestHandler alloc] initWithCGImage:image options:@{}];
    VNRecognizeTextRequest *request = [[VNRecognizeTextRequest alloc] init];
    request.recognitionLevel = VNRequestTextRecognitionLevelAccurate;
    request.usesLanguageCorrection = YES;
    request.recognitionLanguages = @[@"zh-Hans", @"zh-Hant", @"en-US"];
    if (![handler performRequests:@[request] error:error]) return nil;
    NSArray *results = request.results ?: @[];
    NSArray *sorted = [results sortedArrayUsingComparator:^NSComparisonResult(VNRecognizedTextObservation *a, VNRecognizedTextObservation *b) {
        CGFloat ay = a.boundingBox.origin.y, by = b.boundingBox.origin.y;
        if (fabs(ay - by) > 0.01) return ay > by ? NSOrderedAscending : NSOrderedDescending;
        return a.boundingBox.origin.x < b.boundingBox.origin.x ? NSOrderedAscending : NSOrderedDescending;
    }];
    NSMutableArray *lines = [NSMutableArray array];
    for (VNRecognizedTextObservation *obs in sorted) {
        VNRecognizedText *top = [[obs topCandidates:1] firstObject];
        if (top.string.length) [lines addObject:top.string];
    }
    return [lines componentsJoinedByString:@"\n"];
}

static CGImageRef imageFromPage(PDFPage *page, CGFloat scale) {
    NSRect rect = [page boundsForBox:kPDFDisplayBoxMediaBox];
    size_t width = MAX(1, (size_t)(rect.size.width * scale));
    size_t height = MAX(1, (size_t)(rect.size.height * scale));
    CGColorSpaceRef cs = CGColorSpaceCreateWithName(kCGColorSpaceSRGB);
    CGContextRef context = CGBitmapContextCreate(NULL, width, height, 8, 0, cs, (CGBitmapInfo)kCGImageAlphaPremultipliedLast);
    CGColorSpaceRelease(cs);
    if (!context) return NULL;
    CGContextSetFillColorWithColor(context, CGColorGetConstantColor(kCGColorWhite));
    CGContextFillRect(context, CGRectMake(0, 0, width, height));
    CGContextScaleCTM(context, scale, scale);
    [page drawWithBox:kPDFDisplayBoxMediaBox toContext:context];
    CGImageRef image = CGBitmapContextCreateImage(context);
    CGContextRelease(context);
    return image;
}

int main(int argc, const char *argv[]) {
    @autoreleasepool {
        if (argc != 3) { fprintf(stderr, "usage: vision-ocr <input.pdf|image> <output.json>\n"); return 2; }
        NSString *input = [NSString stringWithUTF8String:argv[1]];
        NSString *output = [NSString stringWithUTF8String:argv[2]];
        NSMutableArray *pages = [NSMutableArray array];
        NSError *error = nil;
        PDFDocument *pdf = [[PDFDocument alloc] initWithURL:[NSURL fileURLWithPath:input]];
        if (pdf) {
            if (pdf.isLocked) { fprintf(stderr, "PDF 已加密，无法识别\n"); return 1; }
            for (NSUInteger i = 0; i < pdf.pageCount; i++) {
                PDFPage *page = [pdf pageAtIndex:i];
                CGImageRef image = imageFromPage(page, 2.0);
                if (!image) continue;
                NSString *text = recognize(image, &error);
                CGImageRelease(image);
                if (!text) { fprintf(stderr, "%s\n", error.localizedDescription.UTF8String); return 1; }
                [pages addObject:@{@"page": @(i + 1), @"text": text}];
            }
        } else {
            NSURL *url = [NSURL fileURLWithPath:input];
            CGImageSourceRef source = CGImageSourceCreateWithURL((__bridge CFURLRef)url, NULL);
            if (!source) { fprintf(stderr, "无法读取文件\n"); return 1; }
            CGImageRef image = CGImageSourceCreateImageAtIndex(source, 0, NULL);
            CFRelease(source);
            if (!image) { fprintf(stderr, "无法读取图像\n"); return 1; }
            NSString *text = recognize(image, &error);
            CGImageRelease(image);
            if (!text) { fprintf(stderr, "%s\n", error.localizedDescription.UTF8String); return 1; }
            [pages addObject:@{@"page": @1, @"text": text}];
        }
        NSDictionary *payload = @{@"pages": pages, @"engine": @"macOS Vision"};
        NSData *data = [NSJSONSerialization dataWithJSONObject:payload options:NSJSONWritingPrettyPrinted | NSJSONWritingWithoutEscapingSlashes error:&error];
        if (!data || ![data writeToFile:output atomically:YES]) { fprintf(stderr, "%s\n", error.localizedDescription.UTF8String); return 1; }
    }
    return 0;
}
