import 'dart:io';

import 'package:args/args.dart';
import 'package:gsu_cli/do_stuff.dart';
import 'package:path/path.dart' as path;

Future<void> main(List<String> args) async {
  var parser = ArgParser()
    ..addOption(
      'inputFile',
    )
    ..addOption(
      'templateFile',
    )
    ..addOption(
      'outputFolder',
    );
  var parsed = parser.parse(args);
  var inputPath = parsed['inputFile'];
  var dataStr = await File(inputPath).readAsString();
  print('Reading from $inputPath');
  var templateFilePath = parsed['templateFile'];
  var templateFile = await File(templateFilePath).readAsBytes();
  print('Reading from $inputPath');
  var (name, data) = await computeCapitalGains(dataStr, templateFile);

  var outputFile = path.join(parsed['outputFolder'], '$name.xlsx');
  await File(outputFile).writeAsBytes(data);

  print('Wrote to $outputFile');
}
