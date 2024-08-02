// ignore_for_file: avoid_print

import 'dart:convert';

import 'package:excel/excel.dart';
import 'package:file_saver/file_saver.dart';
import 'package:file_selector/file_selector.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:gsu_cli/do_stuff.dart';

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

    ByteData data = await rootBundle.load('assets/gsu_template.xlsx');
    var bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);

    final (name, resultBytes) = await computeCapitalGains(dataStr, bytes);
    print('Saving file');
    await FileSaver.instance.saveFile(
      name: name,
      bytes: Uint8List.fromList(resultBytes),
    );
  }
}
