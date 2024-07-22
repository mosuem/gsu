// ignore_for_file: avoid_print

import 'dart:convert';

import 'package:excel/excel.dart';
import 'package:file_saver/file_saver.dart';
import 'package:file_selector/file_selector.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

void main() {
  runApp(const MainApp());
}

class MainApp extends StatelessWidget {
  const MainApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: Scaffold(
        body: Center(
          child: Column(
            children: [
              const SelectableText('''
For Morgan Stanley:

Follow the instructions at

https://screenshot.googleplex.com/8rYj3UYcgwKnW8r

Click the button and upload the JSON.

It should download a XLSX. Retrieve the value under TOTAL.

Profit!'''),
              TextButton(
                onPressed: () => readAndWrite(),
                child: const Text('Load and calculate'),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Future<void> readAndWrite() async {
    const XTypeGroup typeGroup = XTypeGroup(
      label: 'jsons',
      extensions: <String>['json'],
    );
    final XFile? file =
        await openFile(acceptedTypeGroups: <XTypeGroup>[typeGroup]);

    String readAsString;
    if (file != null) {
      readAsString = utf8.decode(await file.readAsBytes());
    } else {
      print('Couldnt read file');
      return;
    }
    var decoded = jsonDecode(readAsString) as Map;
    var fromDate = asDate(decoded['FromDate']);
    var toDate = asDate(decoded['ToDate']);
    var transactions = decoded["Transactions"] as List;

    ByteData data = await rootBundle.load('assets/gsu_template.xlsx');
    var bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
    var template = Excel.decodeBytes(bytes);
    var sheetName = template.tables.keys.first;

    var table = template.tables[sheetName]!;
    var j = 6;
    int rowNumber = j;
    print(transactions.length);
    for (var transaction
        in transactions.where((element) => element['Action'] == 'Sale')) {
      for (var detailEl in transaction['TransactionDetails']) {
        var detail = detailEl['Details'] as Map;

        var dataRow = DataRow(
          vesting: Vesting(
            date: asDate(detail['VestDate'] as String),
            fmvUsd: toPrice(detail['VestFairMarketValue'] as String),
            anzahl: double.parse(detail['Shares']),
          ),
          verkauf: Verkauf(
            fmvUsd: toPrice(detail['SalePrice'] as String),
            transaktionskostenUsd: toPrice(transaction['FeesAndCommissions']),
            sellDate: asDate(transaction['Date'] as String),
          ),
        );
        print(dataRow);
        var row = dataRow.toRow(rowNumber);
        for (var i = 0; i < row.length; i++) {
          if (row[i].$1 != null) {
            var cellIndex = CellIndex.indexByColumnRow(
              columnIndex: i,
              rowIndex: rowNumber - 1,
            );
            table.cell(cellIndex).value = row[i].$1;
            table.cell(cellIndex).cellStyle = row[i].$2;
          }
        }
        rowNumber++;
      }
    }
    table.cell(CellIndex.indexByString('Q${rowNumber + 1}')).value =
        const TextCellValue('TOTAL:');
    table.cell(CellIndex.indexByString('P${rowNumber + 2}')).value =
        FormulaCellValue('SUM(P$j:P${rowNumber - 1})');
    table.cell(CellIndex.indexByString('Q${rowNumber + 2}')).value =
        FormulaCellValue('SUM(Q$j:Q${rowNumber - 1})');

    var name = 'GSU_Taxes_${fromDate.fmtDate}_-_${toDate.fmtDate}';
    var resultBytes = template.save(fileName: '$name.xlsx')!;
    print('Saving file');
    await FileSaver.instance.saveFile(
      name: name,
      bytes: Uint8List.fromList(resultBytes),
    );
  }

  double toPrice(String priceStr) => double.parse(priceStr.substring(1));

  DateTime asDate(String dateStr) {
    var split = dateStr.split('/');
    return DateTime(
        int.parse(split[2]), int.parse(split[0]), int.parse(split[1]));
  }
}

class DataRow {
  final Vesting vesting;
  final Verkauf verkauf;

  DataRow({
    required this.vesting,
    required this.verkauf,
  });

  List<(CellValue?, CellStyle)> toRow(int row) {
    var dateStyle = CellStyle(bold: false)
      ..numberFormat = NumFormat.custom(formatCode: 'yyyy-mm-dd');
    var style = CellStyle(bold: false)..numberFormat = NumFormat.standard_2;
    return [
      (DateCellValue.fromDateTime(vesting.date), dateStyle),
      (DoubleCellValue(vesting.fmvUsd), style),
      (
        FormulaCellValue(
            'index(GoogleFinance("CURRENCY:USDEUR", "price", DATE(${verkauf.sellDate.year}, ${verkauf.sellDate.month.toString().padLeft(2, '0')}, ${verkauf.sellDate.day.toString().padLeft(2, '0')})), 2, 2)'),
        style
      ),
      (FormulaCellValue('B$row*C$row'), style),
      (DoubleCellValue(vesting.anzahl), style),
      (FormulaCellValue('D$row*E$row'), style),
      (null, style),
      (DoubleCellValue(verkauf.fmvUsd), style),
      (FormulaCellValue('H$row*C$row'), style),
      (DoubleCellValue(verkauf.transaktionskostenUsd), style),
      (FormulaCellValue('J$row*C$row'), style),
      (FormulaCellValue('I$row*E$row-K$row'), style),
      (null, style),
      (FormulaCellValue('F$row'), style),
      (FormulaCellValue('L$row'), style),
      (FormulaCellValue('O$row-N$row'), style),
      (FormulaCellValue('0.25 * P$row'), style),
    ];
  }

  @override
  String toString() => 'DataRow(vesting: $vesting, verkauf: $verkauf)';
}

class Verkauf {
  final DateTime sellDate;
  final double fmvUsd;
  final double transaktionskostenUsd;

  Verkauf({
    required this.sellDate,
    required this.fmvUsd,
    required this.transaktionskostenUsd,
  });

  @override
  String toString() =>
      'Verkauf(sellDate: $sellDate, fmvUsd: $fmvUsd, transaktionskostenUsd: $transaktionskostenUsd)';
}

class Vesting {
  final DateTime date;
  final double fmvUsd;
  final double anzahl;

  Vesting({
    required this.date,
    required this.fmvUsd,
    required this.anzahl,
  });

  @override
  String toString() => 'Vesting(date: $date, fmvUsd: $fmvUsd, anzahl: $anzahl)';
}

extension on DateTime {
  String get fmtDate => '$year-$month-$day';
}
