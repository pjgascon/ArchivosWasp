<?php
require_once 'vendor/PHPExcel.php';
require_once 'vendor/PHPExcel/IOFactory.php';
require_once 'vendor/PHPExcel/Cell.php';
require_once 'vendor/PHPExcel/Worksheet.php';

$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$archivoConfiguracion = json_decode(file_get_contents('config/palancas.json'), true);
$objPHPExcel = $objReader->load('ENERTEL SURESTE, S.L._Captacion.xlsx');

// Obtengo las familias del archivo de palancas (móvil, ipvpn, sva...)
$arrFamilias = array_keys($archivoConfiguracion['familias']);

foreach ($arrFamilias as $familia) {
    // Aquí iría el procesamiento específico por familia
    procesarFamilia($familia, $archivoConfiguracion, $objPHPExcel);
}

function procesarFamilia($familia, $configuracion, $objPHPExcel)
{
    // Lógica para procesar cada familia
    $filaInicio = $configuracion['familias'][$familia]['inicio'];
    $nombreHoja = $configuracion['familias'][$familia]['hoja'];
    $columnas = array_keys($configuracion['familias'][$familia]['columnas']);

    // Intentamos obtener la hoja
    $sheet = $objPHPExcel->getSheetByName($nombreHoja);

    if ($sheet !== null) {
        $filaFinal = $sheet->getHighestRow();
        for ($fila = $filaInicio; $fila <= $filaFinal; $fila++) {
            foreach ($columnas as $columna) {
                $objColumna = $configuracion['familias'][$familia]['columnas'][$columna];
                $celda = $sheet->getCell($objColumna["pos"] . $fila);
                $tipo = $objColumna["type"];
                $valor = $celda->getValue();

                // Aquí puedes procesar los datos según el tipo
                switch ($tipo) {
                    case 'string':
                        $valor = (string)$valor;
                        break;
                    case 'date':
                        $valor = date("Y-m-d", PHPExcel_Shared_Date::ExcelToPHP($valor));
                        break;
                    case 'bool':
                        $valor = (bool)$valor;
                        break;
                    case 'int':
                        $valor = (int)$valor;
                        break;
                    case 'decimal':
                        $valor = (float)$valor;
                        break;
                }

                // Guardar los datos procesados
                $datos[$columna] = $valor;
            }
            if (strlen($datos['CIF_DISTRIBUIDOR']) > 0) {
                guardarDatos($familia, $datos);
            } else {
                break;
            }
        }
    }
}

function guardarDatos($familia, $datos)
{
    if ($familia == 'movil') {
        $sql = "call callidus_movil_guardar('{$datos['CIF_DISTRIBUIDOR']}',
                                            '{$datos['NOMBRE_DISTRIBUIDOR']}',
                                            '{$datos['SFID']}',
                                            '{$datos['MSISDN']}',
                                            '{$datos['NIF_CLIENTE']}',
                                            '{$datos['FECHA_CICLO']}',
                                            '{$datos['FECHA_ACTIVACION_PROD']}',
                                            '{$datos['FECHA_ACTIVACION_SERVICIO']}',
                                            '{$datos['FECHA_EFECTIVIDAD']}',
                                            '{$datos['TIPO_CLIENTE']}',
                                            '{$datos['SUB_TIPO_CLIENTE']}',
                                            '{$datos['PLAN']}',
                                            '{$datos['TIPO_EVENTO']}',
                                            '{$datos['PORTABILIDAD']}',
                                            '{$datos['TIPOLOGIA_SFID']}',
                                            '{$datos['SVT_VALUE']}',
                                            '{$datos['PROMOCION']}',
                                            '{$datos['FECHA_CARGA1']}',
                                            '{$datos['ORDER_ID']}',
                                            '{$datos['CARTERA_ASIGNADA']}',
                                            '{$datos['COD_PROVEEDOR']}',
                                            '{$datos['CIF_NUEVO']}',
                                            '{$datos['SUM_VALUE']}',
                                            '{$datos['NAME1']}',
                                            '{$datos['VALUE_IND']}',
                                            '{$datos['CICLO']}',
                                            '{$datos['KPI']}',
                                            '{$datos['FAMILIA']}',
                                            '{$datos['CATEGORIA']}',
                                            '{$datos['DESCRIPCION']}',
                                            '{$datos['EQUIVALENCIA']}',
                                            '{$datos['CATEGORIA_CONVERGENCIA']}',
                                            '{$datos['RED_DATOS']}',
                                            '{$datos['CUENTA_INCENTIVO']}',
                                            '{$datos['CLIENTE_CAPTADO']}',
                                            '{$datos['ZONA_SFID']}',
                                            '{$datos['CAPTA_FIDE']}',
                                            '{$datos['CONTEO']}',
                                            '{$datos['CIF_TERRI']}',
                                            '{$datos['ALTA_REFERENCIADA']}');";
        $mysqli = conexion();
        $mysqli->query($sql);
    } elseif ($familia == 'sva') {
        // Guardar datos en la base de datos para la familia sva
        $sql = "call callidus_sva_guardar('{$datos['CIF_DISTRIBUIDOR']}',
                                          '{$datos['NOMBRE_DISTRIBUIDOR']}',
                                          '{$datos['SFID']}',
                                          '{$datos['MSISDN']}',
                                          '{$datos['NIF_CLIENTE']}',
                                          '{$datos['FECHA_CICLO']}',
                                          '{$datos['FECHA_ACTIVACION_SERVICIO']}',
                                          '{$datos['PLAN']}',
                                          '{$datos['TIPO_EVENTO']}',
                                          '{$datos['SUM_VALUE']}',
                                          '{$datos['VALUE_IND']}',
                                          '{$datos['KPI']}',
                                          '{$datos['FAMILIA']}',
                                          '{$datos['DESCRIPCION']}',
                                          '{$datos['EQUIVALENCIA']}',
                                          '{$datos['VALOR_ESTRATEGICO']}');";
        $mysqli = conexion();
        $mysqli->query($sql);
    }
}

function conexion()
{
    static $mysqli = null;

    if ($mysqli instanceof mysqli) {
        // Comprueba si sigue viva
        if (@$mysqli->ping()) {
            return $mysqli;
        }
        // Si no responde, cerramos y reabrimos
        @$mysqli->close();
        $mysqli = null;
    }

    $dbHost = 'desarrolloserver.liberi.es';
    $dbUser = 'root';
    $dbPass = 'Liberi_2024#';
    $dbName = 'connect';
    $dbPort = 3307;

    $mysqli = new mysqli($dbHost, $dbUser, $dbPass, $dbName, $dbPort);

    if ($mysqli->connect_errno) {
        die('Error MySQL (' . $mysqli->connect_errno . '): ' .
            $mysqli->connect_error);
    }

    $mysqli->set_charset('utf8mb4');

    return $mysqli;
}
