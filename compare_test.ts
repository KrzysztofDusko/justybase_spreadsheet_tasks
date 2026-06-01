import { XlsbWriter } from './dist';

async function main() {
  const writer = new XlsbWriter('compare_test.xlsb');

  // 1 - Typy podstawowe
  writer.addSheet('Typy podstawowe');
  writer.writeSheet([
    [42, 3.14159, 'Hello World', new Date(2025, 5, 1), true, null],
    [100, 0.99, 'Test XII', new Date(2025, 11, 24), false, null],
  ], ['Liczba całkowita', 'Zmiennoprzecinkowa', 'Tekst', 'Data', 'Boolean', 'Pusta']);

  // 2 - Finanse
  writer.addSheet('Finanse');
  writer.writeSheet([
    ['Wpłata', 15000.0, null, null, 0.15],
    ['Czynsz', null, '3200.5', null, null],
  ], ['Opis', 'Przychod', 'Wydatek', 'Saldo', 'Procent']);

  // 3 - Oceny
  writer.addSheet('Oceny');
  writer.writeSheet([
    ['Anna Nowak', 5, 4, 6],
    ['Jan Kowalski', 4, 3, 5],
    ['Kasia Wisniewska', 6, 5, 6],
  ], ['Uczen', 'Matematyka', 'Fizyka', 'Informatyka']);

  // 4 - Sprzedaz
  writer.addSheet('Sprzedaz');
  writer.writeSheet([
    ['Q1 2025', 45000, 32000, 28000, 120000],
    ['Q2 2025', 52000, 38000, 31000, 120000],
  ], ['Kwartal', 'Produkt A', 'Produkt B', 'Produkt C', 'Cel']);

  await writer.finalize();
  console.log('OK: compare_test.xlsb');
}

main().catch(console.error);