// ignore_for_file: avoid_print

import 'dart:convert';
import 'package:excel/excel.dart';
import 'package:flutter/services.dart';
import 'package:gsu/main.dart';

Future<(String, List<int>)> computeCapitalGains(String dataStr) async {
  var decoded = jsonDecode(dataStr) as Map;
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
          date: asDate(transaction['Date'] as String),
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
  return (name, resultBytes);
}

double toPrice(String priceStr) => double.parse(priceStr.substring(1));

DateTime asDate(String dateStr) {
  var split = dateStr.split('/');
  return DateTime(
      int.parse(split[2]), int.parse(split[0]), int.parse(split[1]));
}

extension on DateTime {
  String get fmtDate => '$year-$month-$day';
}
