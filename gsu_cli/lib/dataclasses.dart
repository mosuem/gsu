import 'package:excel/excel.dart';

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

  static final tab = <String, String>{
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
