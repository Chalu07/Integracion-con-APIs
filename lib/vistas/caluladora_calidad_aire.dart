import 'package:calculadora_calidad_aire/modelos/ciudad.dart';
import 'package:calculadora_calidad_aire/servicios/calidad_aire_servicio.dart';
import 'package:calculadora_calidad_aire/servicios/ciudad_servicio.dart';
import 'package:flutter/material.dart';

class CalculadoraCalidadAire extends StatefulWidget {
  const CalculadoraCalidadAire({super.key});

  @override
  State<CalculadoraCalidadAire> createState() => _CalculadoraCalidadAireState();
}

class _CalculadoraCalidadAireState extends State<CalculadoraCalidadAire> {
  final _formKey = GlobalKey<FormState>();
  Ciudad? _ciudadSeleccionada;
  final CalidadAireServicio _aireApiServicio = CalidadAireServicio();

  DateTime _fechaSeleccionada = DateTime.now();
  final _txtHorasExposicion = TextEditingController();

  String _resultado = "";
  final CiudadServicio _ciudadesServicio = CiudadServicio();
  List<Ciudad> _ciudades = [];

  // Variable de estado para controlar la carga
  bool _cargando = true;

  Future<void> _calcularIndiceExposicion() async {
    if (_formKey.currentState!.validate()) {
      setState(() {
        _resultado = 'Calculando...';
      });
      final horas = double.tryParse(_txtHorasExposicion.text);
      if (_ciudadSeleccionada != null && horas != null) {
        final pm25 = await _aireApiServicio.obtenerPM25(
          _ciudadSeleccionada!.latitud,
          _ciudadSeleccionada!.longitud,
          _fechaSeleccionada,
        );
        if (pm25 == null || pm25 <= 0) {
          setState(() {
            _resultado =
                "No se pudo obtener el dato de PM2.5. Intenta con otra fecha.";
          });
          return;
        }

        final indice = pm25 * horas;
        String nivel;

        if (indice <= 100) {
          nivel = "Bajo";
        } else if (indice <= 200) {
          nivel = "Moderado";
        } else {
          nivel = "Alto";
        }

        setState(() {
          _resultado =
              "Índice de Exposición: ${indice.toStringAsFixed(2)} - Riesgo: $nivel";
        });
      }
    }
  }

  Future<void> _cargarCiudades() async {
    try {
      final ciudades = await _ciudadesServicio.cargarCiudades();
      setState(() {
        _ciudades = ciudades;
        _cargando = false;
      });
    } catch (e) {
      setState(() {
        _cargando = false;
        _resultado = 'Error al cargar los datos. Revisa el archivo JSON.';
      });
    }
  }

  @override
  void initState() {
    super.initState();
    _cargarCiudades();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: const Color.fromARGB(255, 239, 236, 236),
      appBar: AppBar(
        title: const Text(
          'Calculadora de Calidad del Aire',
          style: TextStyle(
            fontSize: 20,
            fontWeight: FontWeight.w600,
            letterSpacing: 1.2,
            fontFamily: "Roboto",
          ),
        ),
      ),
      body: SingleChildScrollView(
        child: Form(
          key: _formKey,
          child: Column(
            children: [
              if (_cargando)
                const Center(child: CircularProgressIndicator())
              else
                DropdownButtonFormField<Ciudad>(
                  decoration: const InputDecoration(
                    labelText: "Ciudad",
                    hintText: "Elija la ciudad de Colombia",
                    border: OutlineInputBorder(),
                  ),
                  items: _ciudades.map((ciudad) {
                    return DropdownMenuItem<Ciudad>(
                      value: ciudad,
                      child: Text(ciudad.nombre),
                    );
                  }).toList(),
                  onChanged: (Ciudad? ciudad) {
                    setState(() {
                      _ciudadSeleccionada = ciudad;
                    });
                  },
                  validator: (ciudad) =>
                      ciudad == null ? "Debe seleccionar la ciudad" : null,
                ),
              const SizedBox(height: 16),
              Row(
                children: [
                  Expanded(
                    child: TextFormField(
                      decoration: const InputDecoration(
                        labelText: "Fecha Seleccionada",
                      ),
                      readOnly: true,
                      controller: TextEditingController(
                        text: _fechaSeleccionada.toString().substring(0, 10),
                      ),
                      onTap: () async {
                        final valorSeleccionado = await showDatePicker(
                          context: context,
                          firstDate: DateTime(2000),
                          lastDate: DateTime.now(),
                          initialDate: _fechaSeleccionada,
                        );
                        if (valorSeleccionado != null) {
                          setState(() {
                            _fechaSeleccionada = valorSeleccionado;
                          });
                        }
                      },
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 16),
              TextFormField(
                controller: _txtHorasExposicion,
                decoration: const InputDecoration(
                  labelText: "Horas de exposicion al aire libre por dia",
                  border: OutlineInputBorder(),
                ),
                keyboardType: TextInputType.number,
                validator: (valor) {
                  if (valor == null || valor.isEmpty) {
                    return "Ingrese las horas de exposicion";
                  }
                  if (double.tryParse(valor) == null) {
                    return "Ingrese un número válido";
                  }
                  return null;
                },
              ),
              const SizedBox(height: 16),
              ElevatedButton(
                onPressed: _calcularIndiceExposicion,
                child: const Text("Calcular Riesgo de Exposición"),
              ),
              const SizedBox(height: 16),
              Text(_resultado, textAlign: TextAlign.center),
            ],
          ),
        ),
      ),
    );
  }
}
