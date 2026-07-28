// وقتی افزونه لود شد، این تابع اجرا می‌شود
Office.onReady(function(info) {
    if (info.host === Office.HostType.Word) {
        console.log("افزونه داخل ورد اجرا شد!");
    }
});

// تابعی که متن را در ورد درج می‌کند
async function insertText() {
    await Word.run(async (context) => {
        // درج یک پاراگراف جدید در انتهای سند
        context.document.body.insertParagraph("این متن بدون Node.js نوشته شده است!", Word.InsertLocation.end);
        await context.sync();
    });
}