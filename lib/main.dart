// ignore_for_file: avoid_print

import 'dart:convert';

import 'package:excel/excel.dart';
import 'package:file_saver/file_saver.dart';
import 'package:file_selector/file_selector.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:gsu/do_stuff.dart';

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
For Schwab:

Follow the instructions at

https://screenshot.googleplex.com/8rYj3UYcgwKnW8r

Click the button and upload the JSON.

It should download a XLSX. Retrieve the value under TOTAL.

Profit!'''),
              ElevatedButton(
                onPressed: () => readAndWrite(),
                child: const Text('Load and calculate'),
              ),
              const Text('v0.3 mosum@'),
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

    String dataStr;
    if (file != null) {
      dataStr = utf8.decode(await file.readAsBytes());
    } else {
      print('Couldnt read file');
      return;
    }

    final (name, resultBytes) = await computeCapitalGains(dataStr);
    print('Saving file');
    await FileSaver.instance.saveFile(
      name: name,
      bytes: Uint8List.fromList(resultBytes),
    );
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
            'index(GoogleFinance("CURRENCY:USDEUR", "price", DATE(${vesting.date.year}, ${vesting.date.month.toString().padLeft(2, '0')}, ${vesting.date.day.toString().padLeft(2, '0')})), 2, 2)'),
        style
      ),
      (
        FormulaCellValue(
            '${tab['FMV in \$']}$row*${tab['Umrechnungskurs']}$row'),
        style
      ),
      (DoubleCellValue(vesting.anzahl), style),
      (FormulaCellValue('${tab['FMV in €']}$row*${tab['Anzahl']}$row'), style),
      (null, style),
      (DoubleCellValue(verkauf.fmvUsd), style),
      (
        FormulaCellValue(
            '${tab['VFMV in \$']}$row*${tab['VUmrechnungskurs']}$row'),
        style
      ),
      (
        FormulaCellValue(
            'index(GoogleFinance("CURRENCY:USDEUR", "price", DATE(${verkauf.date.year}, ${verkauf.date.month.toString().padLeft(2, '0')}, ${verkauf.date.day.toString().padLeft(2, '0')})), 2, 2)'),
        style
      ),
      (DoubleCellValue(verkauf.transaktionskostenUsd), style),
      (
        FormulaCellValue(
            '${tab['Transaktionskosten in \$']}$row*${tab['VUmrechnungskurs']}$row'),
        style
      ),
      (
        FormulaCellValue(
            '${tab['VFMV in €']}$row*${tab['Anzahl']}$row-${tab['Transaktionskosten in €']}$row'),
        style
      ),
      (null, style),
      (FormulaCellValue('${tab['Total Cost / Gesamte Kosten']}$row'), style),
      (FormulaCellValue('${tab['VTotal Cost / Gesamte Kosten']}$row'), style),
      (
        FormulaCellValue(
            '${tab['Gesamte Kosten Verkauf']}$row-${tab['Gesamte Kosten Vesting']}$row'),
        style
      ),
      (FormulaCellValue('0.25 * ${tab['Gewinn']}$row'), style),
    ];
  }

  final tab = <String, String>{
    'Date': 'A',
    'FMV in \$': 'B',
    'Umrechnungskurs': 'C',
    'FMV in €': 'D',
    'Anzahl': 'E',
    'Total Cost / Gesamte Kosten': 'F',
    'VFMV in \$': 'H',
    'VFMV in €': 'I',
    'VUmrechnungskurs': 'J',
    'Transaktionskosten in \$': 'K',
    'Transaktionskosten in €': 'L',
    'VTotal Cost / Gesamte Kosten': 'M',
    'Gesamte Kosten Vesting': 'O',
    'Gesamte Kosten Verkauf': 'P',
    'Gewinn': 'Q',
    '25%': 'R',
  };

  @override
  String toString() => 'DataRow(vesting: $vesting, verkauf: $verkauf)';
}

class Verkauf {
  final DateTime date;
  final double fmvUsd;
  final double transaktionskostenUsd;

  Verkauf({
    required this.date,
    required this.fmvUsd,
    required this.transaktionskostenUsd,
  });

  @override
  String toString() =>
      'Verkauf(sellDate: $date, fmvUsd: $fmvUsd, transaktionskostenUsd: $transaktionskostenUsd)';
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
