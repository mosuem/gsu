import 'dart:io';

import 'package:args/args.dart';
import 'package:gsu/do_stuff.dart';
import 'package:path/path.dart' as path;

Future<void> main(List<String> args) async {
  var parser = ArgParser()
    ..addOption(
      'inputFile',
    )
    ..addOption(
      'outputFolder',
    );
  var parsed = parser.parse(args);
  var inputPath = parsed['inputFile'];
  print('Reading from $inputPath');
  var dataStr = await File(inputPath).readAsString();
  var (name, data) = await computeCapitalGains(dataStr);

  var outputFile = path.join(parsed['outputFolder'], name);
  await File(outputFile).writeAsBytes(data);

  print('Wrote to $outputFile');
}
