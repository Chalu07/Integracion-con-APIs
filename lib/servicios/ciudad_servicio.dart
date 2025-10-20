import 'dart:convert';
import 'package:calculadora_calidad_aire/modelos/ciudad.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/services.dart';

class CiudadServicio {
  Future<List<Ciudad>> cargarCiudades() async {
    try {
      final String contenidoArchivo = await rootBundle.loadString('assets/datos/capitales_colombia.json');
      final datosJson = jsonDecode(contenidoArchivo) as List;

      return datosJson
        .map((objetoJson) => Ciudad.fromJson(objetoJson) )
        .toList();
    } catch (e) {
      debugPrint ('Error al cargar las ciudades: $e');
      throw Exception('No se pudieron cargar las ciudades: $e');
    }
  }
}