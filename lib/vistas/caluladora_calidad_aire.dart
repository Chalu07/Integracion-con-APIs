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
  final CalidadAireServicio _aireApiServicio = CalidadAireServicio();

  DateTime _fechaSeleccionada = DateTime.now();
  final _txtHorasExposicion = TextEditingController();

  final CiudadServicio _ciudadesServicio = CiudadServicio();
  List<Ciudad> _ciudades = [];
  List<Ciudad> _ciudadesSeleccionadas = [];
  List<ResultadoCalidad> _resultados = [];

  bool _cargando = true;
  bool _calculando = false;

  Future<void> _calcularSimultaneo() async {
    if (!_formKey.currentState!.validate()) return;
    if (_ciudadesSeleccionadas.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text("Seleccione al menos una ciudad")),
      );
      return;
    }

    final horas = double.tryParse(_txtHorasExposicion.text);
    if (horas == null) return;

    setState(() {
      _calculando = true;
      _resultados = [];
    });

    final resultados = await _aireApiServicio.obtenerPM25Simultaneo(
      _ciudadesSeleccionadas,
      _fechaSeleccionada,
      horas,
    );

    setState(() {
      _resultados = resultados;
      _calculando = false;
    });
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
      });
    }
  }

  void _mostrarSelectorCiudades() {
    showDialog(
      context: context,
      builder: (context) {
        final seleccionTemporal = List<Ciudad>.from(_ciudadesSeleccionadas);
        return StatefulBuilder(
          builder: (context, setDialogState) {
            return AlertDialog(
              title: Row(
                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                children: [
                  const Text("Seleccionar ciudades"),
                  TextButton(
                    onPressed: () {
                      if (seleccionTemporal.length == _ciudades.length) {
                        setDialogState(() => seleccionTemporal.clear());
                      } else {
                        setDialogState(() {
                          seleccionTemporal.clear();
                          seleccionTemporal.addAll(_ciudades);
                        });
                      }
                    },
                    child: Text(
                      seleccionTemporal.length == _ciudades.length
                          ? "Ninguna"
                          : "Todas",
                    ),
                  ),
                ],
              ),
              content: SizedBox(
                width: double.maxFinite,
                height: 400,
                child: ListView.builder(
                  itemCount: _ciudades.length,
                  itemBuilder: (context, index) {
                    final ciudad = _ciudades[index];
                    final seleccionada = seleccionTemporal.contains(ciudad);
                    return CheckboxListTile(
                      title: Text(ciudad.nombre),
                      value: seleccionada,
                      onChanged: (valor) {
                        setDialogState(() {
                          if (valor == true) {
                            seleccionTemporal.add(ciudad);
                          } else {
                            seleccionTemporal.remove(ciudad);
                          }
                        });
                      },
                    );
                  },
                ),
              ),
              actions: [
                TextButton(
                  onPressed: () => Navigator.pop(context),
                  child: const Text("Cancelar"),
                ),
                ElevatedButton(
                  onPressed: () {
                    setState(() {
                      _ciudadesSeleccionadas = seleccionTemporal;
                    });
                    Navigator.pop(context);
                  },
                  child: Text("Aceptar (${seleccionTemporal.length})"),
                ),
              ],
            );
          },
        );
      },
    );
  }

  Color _colorPorNivel(String nivel) {
    switch (nivel) {
      case "Bajo":
        return Colors.green;
      case "Moderado":
        return Colors.orange;
      case "Alto":
        return Colors.red;
      default:
        return Colors.grey;
    }
  }

  IconData _iconoPorNivel(String nivel) {
    switch (nivel) {
      case "Bajo":
        return Icons.check_circle;
      case "Moderado":
        return Icons.warning;
      case "Alto":
        return Icons.dangerous;
      default:
        return Icons.help_outline;
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
        padding: const EdgeInsets.all(16),
        child: Form(
          key: _formKey,
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              if (_cargando)
                const Center(child: CircularProgressIndicator())
              else ...[
                InkWell(
                  onTap: _mostrarSelectorCiudades,
                  child: InputDecorator(
                    decoration: const InputDecoration(
                      labelText: "Ciudades",
                      hintText: "Toque para seleccionar ciudades",
                      border: OutlineInputBorder(),
                      suffixIcon: Icon(Icons.arrow_drop_down),
                    ),
                    child: Text(
                      _ciudadesSeleccionadas.isEmpty
                          ? "Seleccione una o más ciudades"
                          : "${_ciudadesSeleccionadas.length} ciudad(es) seleccionada(s)",
                      style: TextStyle(
                        color: _ciudadesSeleccionadas.isEmpty
                            ? Colors.grey
                            : Colors.black87,
                      ),
                    ),
                  ),
                ),
                if (_ciudadesSeleccionadas.isNotEmpty)
                  Padding(
                    padding: const EdgeInsets.only(top: 8),
                    child: Wrap(
                      spacing: 6,
                      runSpacing: 4,
                      children: _ciudadesSeleccionadas.map((ciudad) {
                        return Chip(
                          label: Text(
                            ciudad.nombre,
                            style: const TextStyle(fontSize: 12),
                          ),
                          deleteIcon: const Icon(Icons.close, size: 16),
                          onDeleted: () {
                            setState(() {
                              _ciudadesSeleccionadas.remove(ciudad);
                            });
                          },
                        );
                      }).toList(),
                    ),
                  ),
              ],
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
                  labelText: "Horas de exposición al aire libre por día",
                  border: OutlineInputBorder(),
                ),
                keyboardType: TextInputType.number,
                validator: (valor) {
                  if (valor == null || valor.isEmpty) {
                    return "Ingrese las horas de exposición";
                  }
                  if (double.tryParse(valor) == null) {
                    return "Ingrese un número válido";
                  }
                  return null;
                },
              ),
              const SizedBox(height: 16),
              ElevatedButton.icon(
                onPressed: _calculando ? null : _calcularSimultaneo,
                icon: _calculando
                    ? const SizedBox(
                        width: 18,
                        height: 18,
                        child: CircularProgressIndicator(
                          strokeWidth: 2,
                          color: Colors.white,
                        ),
                      )
                    : const Icon(Icons.compare_arrows),
                label: Text(
                  _calculando
                      ? "Consultando ${_ciudadesSeleccionadas.length} ciudad(es)..."
                      : "Calcular Simultáneamente",
                ),
              ),
              const SizedBox(height: 24),
              if (_resultados.isNotEmpty) ...[
                Text(
                  "Resultados - ${_resultados.length} ciudad(es)",
                  style: const TextStyle(
                    fontSize: 18,
                    fontWeight: FontWeight.bold,
                  ),
                ),
                const SizedBox(height: 12),
                ..._resultados.map((r) {
                  return Card(
                    margin: const EdgeInsets.only(bottom: 10),
                    elevation: 2,
                    child: ListTile(
                      leading: Icon(
                        _iconoPorNivel(r.nivel),
                        color: _colorPorNivel(r.nivel),
                        size: 32,
                      ),
                      title: Text(
                        r.ciudad.nombre,
                        style: const TextStyle(fontWeight: FontWeight.w600),
                      ),
                      subtitle: r.indice != null
                          ? Text(
                              "PM2.5: ${r.pm25!.toStringAsFixed(2)} · "
                              "Índice: ${r.indice!.toStringAsFixed(2)}",
                            )
                          : const Text("Sin datos disponibles"),
                      trailing: Container(
                        padding: const EdgeInsets.symmetric(
                          horizontal: 10,
                          vertical: 4,
                        ),
                        decoration: BoxDecoration(
                          color: _colorPorNivel(r.nivel).withAlpha(30),
                          borderRadius: BorderRadius.circular(12),
                          border: Border.all(
                            color: _colorPorNivel(r.nivel),
                          ),
                        ),
                        child: Text(
                          r.nivel,
                          style: TextStyle(
                            color: _colorPorNivel(r.nivel),
                            fontWeight: FontWeight.bold,
                          ),
                        ),
                      ),
                    ),
                  );
                }),
              ],
            ],
          ),
        ),
      ),
    );
  }
}
