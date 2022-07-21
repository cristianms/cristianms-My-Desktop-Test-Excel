// ignore_for_file: avoid_print, import_of_legacy_library_into_null_safe, depend_on_referenced_packages

import 'dart:io';

import 'package:desktop_drop/desktop_drop.dart';
import 'package:desktop_window/desktop_window.dart';
import 'package:excel/excel.dart';
import 'package:path/path.dart';
import 'package:flutter/material.dart';
import 'package:flutter_emoji/flutter_emoji.dart';
import 'package:cross_file/cross_file.dart';

void main() => runApp(const MyApp());

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'XLS Test',
      theme: ThemeData(primarySwatch: Colors.blue, brightness: Brightness.dark),
      home: const MyHomePage(),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key});

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  final List<XFile> list = [];

  bool dragging = false;

  @override
  Widget build(BuildContext context) {
    var parser = EmojiParser();

    return FutureBuilder(
        future: DesktopWindow.setWindowSize(const Size(400, 400)),
        builder: (context, snapshot) {
          if (snapshot.connectionState == ConnectionState.waiting) {
            return const Center(
              child: CircularProgressIndicator(),
            );
          }
          return Scaffold(
            appBar: AppBar(
              title: const Text('XLS Test'),
            ),
            body: DropTarget(
              onDragDone: (detail) {
                setState(() {
                  dragging = false;
                  list.addAll(detail.files);
                  print(list.length);
                });
              },
              onDragEntered: (detail) {
                setState(() {
                  dragging = true;
                });
              },
              onDragExited: (detail) {
                setState(() {
                  dragging = false;
                });
              },
              child: Center(
                child: Column(
                  mainAxisAlignment: MainAxisAlignment.center,
                  children: <Widget>[
                    const Text('Arraste o arquivo gerado no aplicativo "Ponto Fácil" para cá!'),
                    Text(
                      parser.get('open_file_folder').code,
                      style: TextStyle(
                        fontSize: dragging ? 60 : 40,
                      ),
                    ),
                    if (list.isNotEmpty) ...[
                      const Divider(),
                      ...list.map((e) => Text(e.path)).toList(),
                      Padding(
                        padding: const EdgeInsets.only(top: 16, bottom: 16),
                        child: GestureDetector(
                          child: const Text('Limpar arquivos', style: TextStyle(fontWeight: FontWeight.bold)),
                          onTap: () {
                            setState(() {
                              list.clear();
                            });
                          },
                        ),
                      ),
                    ]
                  ],
                ),
              ),
            ),
            floatingActionButton: FloatingActionButton(
              onPressed: _geraXls,
              tooltip: 'Gera arquivo',
              child: const Icon(Icons.file_copy),
            ),
          );
        });
  }

  void _geraXls() {
    if (list.isNotEmpty) {
      for (var arquivo in list) {
        // var file = "/Users/kawal/Desktop/form.xlsx";
        var file = arquivo.path;
        var bytes = File(file).readAsBytesSync();

        // var excel = Excel.createExcel();
        // // or
        // //var excel = Excel.decodeBytes(bytes);
        var excel = Excel.decodeBytes(bytes);
        // for (var table in excel.tables.keys) {
        //   print(table);
        //   print(excel.tables[table]!.maxCols);
        //   print(excel.tables[table]!.maxRows);
        //   for (var row in excel.tables[table]!.rows) {
        //     print("$row");
        //   }
        // }
        for (var table in excel.tables.keys) {
          print(table);
          print(excel.tables[table]!.maxCols);
          print(excel.tables[table]!.maxRows);
          for (var row in excel.tables[table]!.rows) {
            print("$row");
          }
        }

        CellStyle cellStyle = CellStyle(
          bold: true,
          italic: true,
          fontFamily: getFontFamily(FontFamily.Comic_Sans_MS),
        );

        var sheet = excel['Sheet1'];

        var cell = sheet.cell(CellIndex.indexByString("A1"));
        cell.value = "Heya How are you I am fine ok goood night";
        cell.cellStyle = cellStyle;

        var cell2 = sheet.cell(CellIndex.indexByString('E5'));
        cell2.value = 'Heya How night';
        cell2.cellStyle = cellStyle;

        // /// printing cell-type
        // print('CellType: ${cell.cellType.toString()}');

        // /// Iterating and changing values to desired type
        // for (int row = 0; row < sheet.maxRows; row++) {
        //   sheet.row(row).forEach(
        //     (cell) {
        //       // var val = cell.value; //  Value stored in the particular cell
        //       cell?.value = ' My custom Value ';
        //     },
        //   );
        // }

        // /// appending rows
        // List<List<String>> list = List.generate(6000, (index) => List.generate(20, (index1) => '$index $index1'));

        // Stopwatch stopwatch = Stopwatch()..start();
        // for (var row in list) {
        //   sheet.appendRow(row);
        // }

        // print('doSomething() executed in ${stopwatch.elapsed}');

        // sheet.appendRow([8]);
        // // excel.setDefaultSheet(sheet.sheetName).then(
        // //   (isSet) {
        // //     // isSet is bool which tells that whether the setting of default sheet is successful or not.
        // //     if (isSet) {
        // //       print("${sheet.sheetName} is set to default sheet.");
        // //     } else {
        // //       print("Unable to set ${sheet.sheetName} to default sheet.");
        // //     }
        // //   },
        // // );
        // var isSet = excel.setDefaultSheet(sheet.sheetName);
        // if (isSet) {
        //   print("${sheet.sheetName} is set to default sheet.");
        // } else {
        //   print("Unable to set ${sheet.sheetName} to default sheet.");
        // }

        // Saving the file
        // String outputFile = "/Users/kawal/Desktop/form1.xlsx";
        String outputFile = "C:/Users/Cristian/Desktop/form1.xlsx";
        var bytesListValues = excel.encode();
        File(join(outputFile))
          ..createSync(recursive: true)
          ..writeAsBytesSync(bytesListValues!);
      }
    }
  }
}
