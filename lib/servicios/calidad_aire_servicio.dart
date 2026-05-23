import "dart:convert";
import "package:calculadora_calidad_aire/modelos/ciudad.dart";
import "package:flutter/foundation.dart";
import "package:http/http.dart" as http;

class ResultadoCalidad {
  final Ciudad ciudad;
  final double? pm25;
  final double horas;
  final double? indice;
  final String nivel;

  ResultadoCalidad({
    required this.ciudad,
    required this.pm25,
    required this.horas,
    required this.indice,
    required this.nivel,
  });
}

class CalidadAireServicio {
  static const urlBase = "https://air-quality-api.open-meteo.com/v1/air-quality";

  Future<double?> obtenerPM25(
    double latitud,
    double longitud,
    DateTime fecha,
  ) async {
    final String fechaStr = fecha.toIso8601String().substring(0, 10);

    final url = Uri.parse(
      "$urlBase?latitude=$latitud"
      "&longitude=$longitud"
      "&hourly=pm2_5"
      "&start_date=$fechaStr"
      "&end_date=$fechaStr"
    );

    try {
      final respuesta = await http.get(url);

      if (respuesta.statusCode == 200) {
        final datos = jsonDecode(respuesta.body);
        final List<dynamic>? pm25Horas = datos['hourly']?['pm2_5'];

        if (pm25Horas != null && pm25Horas.isNotEmpty) {
          final promedio = pm25Horas
            .map((valor) => valor is num ? valor.toDouble() : 0.0)
            .reduce((a, b) => a + b) / pm25Horas.length;

          return promedio;
        }
      } else {
        debugPrint("Error en la respuesta: ${respuesta.statusCode}");
      }
    } catch (e) {
      debugPrint("Error al obtener datos de la API: $e");
    }

    return null;
  }

  Future<List<ResultadoCalidad>> obtenerPM25Simultaneo(
    List<Ciudad> ciudades,
    DateTime fecha,
    double horas,
  ) async {
    final futures = ciudades.map((ciudad) async {
      final pm25 = await obtenerPM25(ciudad.latitud, ciudad.longitud, fecha);
      double? indice;
      String nivel = "Sin datos";

      if (pm25 != null && pm25 > 0) {
        indice = pm25 * horas;
        if (indice <= 100) {
          nivel = "Bajo";
        } else if (indice <= 200) {
          nivel = "Moderado";
        } else {
          nivel = "Alto";
        }
      }

      return ResultadoCalidad(
        ciudad: ciudad,
        pm25: pm25,
        horas: horas,
        indice: indice,
        nivel: nivel,
      );
    }).toList();

    return Future.wait(futures);
  }
}
