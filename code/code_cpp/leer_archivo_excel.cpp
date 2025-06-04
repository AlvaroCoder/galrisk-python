#include <OpenXLSX/OpenXLSX.hpp>
#include <iostream>
#include <string>

int main(int argc, char* argv[]) {
    if (argc != 4) {
        std::cerr << "Uso: ./leer_celda archivo.xlsx hoja celda\n";
        return 1;
    }

    std::string ruta = argv[1];
    std::string hoja = argv[2];
    std::string celda = argv[3];

    try {
        OpenXLSX::XLDocument doc;
        doc.open(ruta);
        auto wks = doc.workbook().worksheet(hoja);

        auto valor = wks.cell(celda).value();

        if (valor.isEmpty()) {
            std::cout << "None\n";
        } else if (valor.type() == OpenXLSX::XLValueType::Integer) {
            std::cout << valor.get<int>() << "\n";
        } else if (valor.type() == OpenXLSX::XLValueType::Float) {
            std::cout << valor.get<double>() << "\n";
        } else if (valor.type() == OpenXLSX::XLValueType::String) {
            std::cout << valor.get<std::string>() << "\n";
        } else {
            std::cout << "Unsupported value type\n";
        }

        doc.close();
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << "\n";
        return 1;
    }

    return 0;
}