import { XlsbWriter, XlsxWriter, F } from './dist';
import * as path from 'path';

const outputDir = __dirname;

function sheets(writer: XlsbWriter | XlsxWriter) {
  // 1
  writer.addSheet('Typy podstawowe');
  writer.writeSheet([
    [42, 3.14159, 'Hello World', new Date(2025, 5, 1), true, null],
    [100, 0.99, 'Test XII', new Date(2025, 11, 24), false, undefined],
  ], ['Liczba całkowita', 'Zmiennoprzecinkowa', 'Tekst', 'Data', 'Boolean', 'Pusta']);

  // 2
  writer.addSheet('Finanse');
  writer.writeSheet([
    ['Wpłata', { value: 15000.0, format: F.CURRENCY_PLN }, null, null, { value: 0.15, format: F.PERCENTAGE }],
    ['Czynsz', null, { value: 3200.5, format: F.CURRENCY_PLN }, null, null],
    ['Zakupy', null, { value: 1240.99, format: F.CURRENCY_PLN }, null, null],
  ], ['Opis', 'Przychód', 'Wydatek', 'Saldo', 'Procent']);

  // 3
  writer.addSheet('Daty i czas');
  writer.writeSheet([
    [{ value: new Date(2025, 0, 1), format: F.DATE_ISO }, '08:00:00', { value: new Date(2025, 0, 1, 8, 0), format: F.DATETIME_ISO }, 'Środa'],
    [{ value: new Date(2025, 5, 15), format: F.DATE_ISO }, '14:30:00', { value: new Date(2025, 5, 15, 14, 30), format: F.DATETIME_ISO }, 'Niedziela'],
    [{ value: new Date(2025, 11, 25), format: F.DATE_ISO }, '23:59:59', { value: new Date(2025, 11, 25, 23, 59, 59), format: F.DATETIME_ISO }, 'Czwartek'],
  ], ['Data', 'Czas', 'Data i czas', 'Dzień tygodnia']);

  // 4
  writer.addSheet('Naukowe');
  writer.writeSheet([
    ['Stała Plancka', { value: 6.62607015e-34, format: F.SCIENTIFIC }, '6.626×10⁻³⁴', 'J·s'],
    ['Prędkość światła', { value: 299792458, format: F.THOUSANDS_SEP }, '2.998×10⁸', 'm/s'],
    ['Stała grawitacji', { value: 6.6743e-11, format: F.SCIENTIFIC }, '6.674×10⁻¹¹', 'm³·kg⁻¹·s⁻²'],
    ['Liczba Avogadro', { value: 6.02214076e23, format: F.SCIENTIFIC }, '6.022×10²³', 'mol⁻¹'],
  ], ['Parametr', 'Wartość', 'Notacja', 'Jednostka']);

  // 5
  writer.addSheet('Geograficzne');
  writer.writeSheet([
    ['Warszawa', 'Polska', 52.2297, 21.0122, 'UTC+1', 1.79],
    ['Kraków', 'Polska', 50.0647, 19.945, 'UTC+1', 0.77],
    ['Gdańsk', 'Polska', 54.352, 18.6466, 'UTC+1', 0.47],
    ['Wrocław', 'Polska', 51.1079, 17.0385, 'UTC+1', 0.64],
    ['Poznań', 'Polska', 52.4064, 16.9252, 'UTC+1', 0.54],
  ], ['Miasto', 'Kraj', 'Szerokość', 'Długość', 'Strefa', 'Populacja (mln)']);

  // 6
  writer.addSheet('Magazyn');
  writer.writeSheet([
    ['P-001', 'Laptop Dell', 'Elektronika', { value: 4500, format: F.CURRENCY_PLN }, '23%', 15],
    ['P-002', 'Monitor LG', 'Elektronika', { value: 1200, format: F.CURRENCY_PLN }, '23%', 30],
    ['P-003', 'Klawiatura', 'Akcesoria', { value: 89.99, format: F.CURRENCY_PLN }, '23%', 120],
    ['P-004', 'Mysz Logitech', 'Akcesoria', { value: 149.99, format: F.CURRENCY_PLN }, '23%', 85],
    ['P-005', 'Biurko', 'Meble', { value: 890, format: F.CURRENCY_PLN }, '8%', 10],
  ], ['ID', 'Nazwa', 'Kategoria', 'Cena netto', 'VAT', 'Stan']);

  // 7
  writer.addSheet('Oceny');
  writer.writeSheet([
    ['Anna Nowak', 5, 4, 6, 5, 5],
    ['Jan Kowalski', 4, 3, 5, 4, 4],
    ['Kasia Wiśniewska', 6, 5, 6, 5, 6],
    ['Tomasz Lewandowski', 3, 4, 4, 3, 4],
    ['Marta Zielińska', 5, 5, 5, 6, 5],
  ], ['Uczeń', 'Matematyka', 'Fizyka', 'Informatyka', 'Polski', 'Angielski']);

  // 8
  writer.addSheet('Różne typy');
  writer.writeSheet([
    ['Integer', 1, 2, 3, 'liczby całkowite'],
    ['Float', 0.1, 0.01, 0.001, 'liczby zmiennoprzecinkowe'],
    ['String', 'Ala', 'ma', 'kota', 'ciągi znaków'],
    ['Date', new Date(2025, 0, 1), new Date(2025, 5, 1), new Date(2025, 11, 31), 'daty'],
    ['Boolean', true, false, true, 'wartości logiczne'],
    ['Null', null, null, null, 'puste komórki'],
  ], ['Typ', 'Wartość 1', 'Wartość 2', 'Wartość 3', 'Opis']);

  // 9
  writer.addSheet('Pracownicy');
  writer.writeSheet([
    [1, 'Adam', 'Mickiewicz', 'IT', 'Developer', new Date(2020, 2, 1), { value: 12000, format: F.CURRENCY_PLN }],
    [2, 'Maria', 'Skłodowska', 'R&D', 'Naukowiec', new Date(2019, 6, 15), { value: 15000, format: F.CURRENCY_PLN }],
    [3, 'Mikołaj', 'Kopernik', 'R&D', 'Astronom', new Date(2018, 0, 1), { value: 18000, format: F.CURRENCY_PLN }],
    [4, 'Fryderyk', 'Chopin', 'HR', 'Specjalista', new Date(2021, 10, 1), { value: 9000, format: F.CURRENCY_PLN }],
    [5, 'Wisława', 'Szymborska', 'Marketing', 'Kreatywna', new Date(2022, 4, 1), { value: 11000, format: F.CURRENCY_PLN }],
  ], ['ID', 'Imię', 'Nazwisko', 'Dział', 'Stanowisko', 'Data zatrudnienia', 'Pensja']);

  // 10
  writer.addSheet('Sprzedaż');
  writer.writeSheet([
    ['Q1 2025', 45000, 32000, 28000, 120000],
    ['Q2 2025', 52000, 38000, 31000, 120000],
    ['Q3 2025', 48000, 35000, 29000, 120000],
    ['Q4 2025', 61000, 42000, 35000, 120000],
  ], ['Kwartał', 'Produkt A', 'Produkt B', 'Produkt C', 'Cel']);

  // 11
  writer.addSheet('Szablon');
  const tpl: any[] = [];
  for (let i = 1; i <= 10; i++) tpl.push([i, '', 0, 0.0]);
  writer.writeSheet(tpl, ['Lp.', 'Nazwa', 'Ilość', 'Cena']);

  // 12
  writer.addSheet('Notatki');
  writer.writeSheet([
    [1, 'Ważne spotkanie', 'Spotkanie z klientem dot. nowego projektu. Omówienie wymagań i harmonogramu.', 'AD', new Date(2025, 5, 1), 'Wysoki'],
    [2, 'Kodowanie', 'Implementacja modułu eksportu danych do formatu XLSX i XLSB. Testy jednostkowe wymagane.', 'JK', new Date(2025, 5, 2), 'Średni'],
    [3, 'Błąd krytyczny', 'Wystąpił wyjątek przy parsowaniu dużych plików CSV. Należy zoptymalizować parser.', 'AN', new Date(2025, 5, 3), 'Krytyczny'],
    [4, 'Wdrożenie', 'Przygotowanie środowiska produkcyjnego. Konfiguracja serwerów i bazy danych.', 'ML', new Date(2025, 5, 4), 'Niski'],
  ], ['ID', 'Tytuł', 'Treść', 'Autor', 'Data', 'Priorytet']);

  // 13
  writer.addSheet('Procenty');
  const total = 450000 + 320000 + 180000;
  writer.writeSheet([
    ['Sprzedaż online', 450000, { value: 450000 / total, format: F.PERCENTAGE }],
    ['Sprzedaż stacjonarna', 320000, { value: 320000 / total, format: F.PERCENTAGE }],
    ['Eksport', 180000, { value: 180000 / total, format: F.PERCENTAGE }],
  ], ['Kategoria', 'Wartość', 'Udział %']);

  // 14
  writer.addSheet('Flagi');
  writer.writeSheet([
    [1, 'user1', true, false, false],
    [2, 'user2', true, false, true],
    [3, 'user3', false, false, false],
    [4, 'admin1', true, true, false],
    [5, 'mod1', true, false, true],
  ], ['ID', 'Nazwa', 'Aktywny', 'Admin', 'Zablokowany']);

  // 15
  writer.addSheet('Spis treści');
  writer.writeSheet([
    [1, 'Typy podstawowe', 'int, float, string, date, bool, null', 2, 6],
    [2, 'Finanse', 'waluty i procenty z F.CURRENCY_PLN / F.PERCENTAGE', 3, 5],
    [3, 'Daty i czas', 'date, time, datetime ISO', 3, 4],
    [4, 'Naukowe', 'notacja naukowa F.SCIENTIFIC', 4, 4],
    [5, 'Geograficzne', 'współrzędne miast', 5, 6],
    [6, 'Magazyn', 'ceny F.CURRENCY_PLN, stany magazynowe', 5, 6],
    [7, 'Oceny', 'oceny szkolne 1-6', 5, 6],
    [8, 'Różne typy', 'miks: int, float, string, date, bool, null', 6, 5],
    [9, 'Pracownicy', 'dane osobowe, daty, pensje F.CURRENCY_PLN', 5, 7],
    [10, 'Sprzedaż', 'sprzedaż kwartalna', 4, 5],
    [11, 'Szablon', 'pusty szablon 10 wierszy', 10, 4],
    [12, 'Notatki', 'długie teksty, daty, priorytety', 4, 6],
    [13, 'Procenty', 'udziały procentowe F.PERCENTAGE', 3, 3],
    [14, 'Flagi', 'wartości logiczne boolean', 5, 5],
    [15, 'Spis treści', 'niniejsza zakładka', 15, 5],
  ], ['Nr', 'Nazwa zakładki', 'Opis', 'Wierszy', 'Kolumn']);
}

async function main() {
  const xlsxPath = path.join(outputDir, 'data_formats.xlsx');
  const xlsxWriter = new XlsxWriter(xlsxPath);
  sheets(xlsxWriter);
  await xlsxWriter.finalize();
  console.log(`OK: ${xlsxPath}`);

  const xlsbPath = path.join(outputDir, 'data_formats.xlsb');
  const xlsbWriter = new XlsbWriter(xlsbPath);
  sheets(xlsbWriter);
  await xlsbWriter.finalize();
  console.log(`OK: ${xlsbPath}`);

  console.log(`\nRazem: 15 zakladek w kazdym pliku.`);
}

main().catch(console.error);
