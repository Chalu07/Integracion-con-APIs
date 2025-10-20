import "dart:convert";
import "package:flutter/foundation.dart";
import "package:http/http.dart" as http;

class CalidadAireServicio {
  static const urlBase = "https://air-quality-api.open-meteo.com/v1/air-quality";

  Future<double?> obtenerPM25(
    double latitud,
    double longitud,
    DateTime fecha,
  ) async {
    final String fechaStr = fecha.toIso8601String().substring(0, 10);

    final url = Uri.parse(
      "$urlBase?latitude=$latitud""&longitude=$longitud""&hourly=pm2_5""&start_date=$fechaStr""&end_date=$fechaStr"
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
}
