// ignore_for_file: avoid_print, import_of_legacy_library_into_null_safe, depend_on_referenced_packages

import 'dart:convert';

import 'package:cross_file/cross_file.dart';
import 'package:desktop_window/desktop_window.dart';
import 'package:excel/excel.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter_emoji/flutter_emoji.dart';
import 'package:path/path.dart';

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
  final List<XFile> listaArquivosLeitura = [];
  String caminhoDestino = '';
  List<RegistroDiario> listaRegistrosDiarios = [];

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
          body: Center(
            child: Column(
              mainAxisAlignment: MainAxisAlignment.center,
              children: <Widget>[
                const Text('Clique para fazer a leitura do arquivo de backup no seu Google Drive'),
                // Ícone de pasta
                Text(
                  parser.get('open_file_folder').code,
                  style: const TextStyle(fontSize: 50),
                ),
              ],
            ),
          ),
          floatingActionButton: FloatingActionButton(
            onPressed: () {},
            tooltip: 'Gera arquivo',
            child: const Icon(Icons.file_copy),
          ),
        );
      },
    );
  }

  // String _letraCorrespondente(int n) {
  //   switch (n) {
  //     case 0:
  //       return 'A';
  //     case 1:
  //       return 'B';
  //     case 2:
  //       return 'C';
  //     case 3:
  //       return 'D';
  //     case 4:
  //       return 'E';
  //     case 5:
  //       return 'F';
  //     case 6:
  //       return 'G';
  //     case 7:
  //       return 'H';
  //     case 8:
  //       return 'I';
  //     case 9:
  //       return 'J';
  //     case 10:
  //       return 'K';
  //     case 11:
  //       return 'L';
  //     case 12:
  //       return 'M';
  //     case 13:
  //       return 'N';
  //     case 14:
  //       return 'O';
  //     case 15:
  //       return 'P';
  //     case 16:
  //       return 'Q';
  //     case 17:
  //       return 'R';
  //     case 18:
  //       return 'S';
  //     case 19:
  //       return 'T';
  //     case 20:
  //       return 'U';
  //     case 21:
  //       return 'V';
  //     case 22:
  //       return 'X';
  //     case 23:
  //       return 'Z';
  //     case 24:
  //       return 'W';
  //     case 25:
  //       return 'Y';
  //     case 26:
  //       return 'AA';
  //     case 27:
  //       return 'AB';
  //     case 28:
  //       return 'AC';
  //     case 29:
  //       return 'AD';
  //     case 30:
  //       return 'AE';
  //   }
  //   return '';
  // }

  // void _geraXls() {
  //   if (listaArquivosLeitura.isNotEmpty) {
  //     for (var i = 0; i < listaArquivosLeitura.length; i++) {
  //       final caminhoArquivoLeitura = listaArquivosLeitura[i].path;
  //       final bytesArquivoLeitura = File(caminhoArquivoLeitura).readAsBytesSync();
  //       // Inicializa obj Excel para leitura
  //       final excelLeitura = Excel.decodeBytes(bytesArquivoLeitura);
  //       // Inicializa obj Excel para escrita
  //       final excelCriacao = Excel.createExcel();
  //       final aba = excelCriacao['Sheet1'];
  //       // Percorre cada aba do Excel
  //       for (var abaLeituraAtual in excelLeitura.tables.keys) {
  //         var identificadorDataAtual = '';
  //         // Percorre cada aba do Excel
  //         for (var k = 0; k < excelLeitura.tables[abaLeituraAtual]!.rows.length; k++) {
  //           if (k >= 8) {
  //             final linha = excelLeitura.tables[abaLeituraAtual]!.rows[k];
  //             if (linha[1] != null &&
  //                 linha[1].toString().contains('Data(') &&
  //                 !linha[0].toString().contains('Nenhu') &&
  //                 !linha[1].toString().contains('Nenhu') &&
  //                 !linha[2].toString().contains('Nenhu')) {
  //               identificadorDataAtual = linha[1].toString();
  //               final registroDiario = RegistroDiario(identificador: identificadorDataAtual);
  //               registroDiario.addHorario(linha[2].toString());
  //               listaRegistrosDiarios.add(registroDiario);
  //             } else if (linha[1] == null && linha[2].toString().contains('Data(') && !linha[2].toString().contains('Resumo')) {
  //               listaRegistrosDiarios.firstWhere((element) => element.identificador == identificadorDataAtual).addHorario(linha[2].toString());
  //             }
  //           }
  //         }
  //       }

  //       var indexLinhaDia = 1; // Um dia por linha
  //       for (final dia in listaRegistrosDiarios) {
  //         final horariosPorDia = dia.horarios;
  //         aba.cell(CellIndex.indexByString('A$indexLinhaDia')).value = dia.identificador;
  //         for (var indexHorarioColuna = 0; indexHorarioColuna < horariosPorDia.length; indexHorarioColuna++) {
  //           final letra = _letraCorrespondente(indexHorarioColuna + 1);
  //           aba.cell(CellIndex.indexByString('$letra$indexLinhaDia')).value = horariosPorDia[indexHorarioColuna];
  //         }
  //         indexLinhaDia++;
  //       }

  //       final caminhoArquivoEscrita = '$caminhoDestino/form_${i + 1}.xlsx';
  //       final bytesArquivoEscrita = excelCriacao.encode();
  //       File(join(caminhoArquivoEscrita))
  //         ..createSync(recursive: true)
  //         ..writeAsBytesSync(bytesArquivoEscrita!);
  //     }
  //   }
  // }

// /// Data(07:56 - 11:28, 2, 10, null, Extrato)
//   String _trataHorario(String strXls) {
//     return strXls.substring(0);
//   }
}

class RegistroDiario {
  String identificador;
  List<String> horarios = <String>[];
  RegistroDiario({
    required this.identificador,
  });

  // Data(07:56 - 11:28, 2, 10, null, Extrato)
  void addHorario(String horario) {
    var h1 = horario.substring(5, 10);
    horarios.add(h1);
    var h2 = horario.substring(13, 18);
    horarios.add(h2);
  }

  Map<String, dynamic> toMap() {
    return {
      'identificador': identificador,
      'horarios': horarios,
    };
  }

  factory RegistroDiario.fromMap(Map<String, dynamic> map) {
    return RegistroDiario(
      identificador: map['identificador'] ?? '',
    );
  }

  String toJson() => json.encode(toMap());

  factory RegistroDiario.fromJson(String source) => RegistroDiario.fromMap(json.decode(source));

  @override
  String toString() => 'RegistroDiario(identificador: $identificador, horarios: $horarios)';

  @override
  bool operator ==(Object other) {
    if (identical(this, other)) return true;

    return other is RegistroDiario && other.identificador == identificador && listEquals(other.horarios, horarios);
  }

  @override
  int get hashCode => identificador.hashCode ^ horarios.hashCode;
}
