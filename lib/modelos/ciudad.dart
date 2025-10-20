class Ciudad{
  final String nombre;
  final double latitud;
  final double longitud;

  Ciudad({
    required this.nombre,
    required this.latitud,
    required this.longitud,
  }); 

  factory Ciudad.fromJson(Map<String, dynamic> objetoJson){
    return Ciudad(
      nombre: objetoJson["nombre"],
      latitud: objetoJson["latitud"],
      longitud: objetoJson["longitud"],
    );
  }
}