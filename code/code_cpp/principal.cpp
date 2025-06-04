#include <OpenXLSX/OpenXLSX.hpp>

int main() {
    using namespace OpenXLSX;

    XLDocument doc;
    doc.create("nuevo_archivo.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    wks.cell("A1").value() = "Hola";
    wks.cell("A2").value() = 123.45;

    doc.save();
    doc.close();
    return 0;
}