const express = require("express");
const router = express.Router();
const Excel = require("exceljs");
const axios = require("axios");
const _ = require("underscore");
const lo = require("lodash");
const dateFormat = require("dateformat");
const moment = require("moment-timezone");
const fs = require("fs");
const date = dateFormat(new Date(), "dddmmmddyyyyHHMMss");
const salida = moment()
  .tz("America/Lima")
  .format("HH:mm:ss");

const configs = require("../configs/configs");
const END_POINT = configs.endPoint;
const SUB_END_POINT = configs.subEndPoint;

const TEMPLATE = "./template/eidoCashierReport.xlsx";
let cierre = "";

const A = 1;
const B = 2;
const C = 3;
const D = 4;
const E = 5;
const F = 6;
const G = 7;
const H = 8;
const I = 9;
const J = 10;
const K = 11;
const L = 12;
const M = 13;
const N = 14;

const getCierre = cierre => {
  return new Promise((resolve, reject) => {
    axios
      .get(`${END_POINT}cierres/${cierre}`, {
        headers: { "Access-Control-Allow-Origin": "*" },
        responseType: "json"
      })
      .then(response => {
        resolve(response.data);
      })
      .catch(error => {
        reject(error.message);
      });
  });
};

const calculateTotals = cierre => {
  return new Promise((resolve, reject) => {
    let fondo = cierre.fondo;
    let pax = fondo.pax;
    let totalFondo = fondo.totalFondo;
    let desvio = cierre.desvio;
    let detalleAuto = fondo.detalleAuto ? fondo.detalleAuto : [];

    let moneda = _.omit(
      fondo,
      "detalleCierre",
      "detalleAuto",
      "totalFondo",
      "turno",
      "_id",
      "pax",
      "__v"
    );

    let empleado = {
      nombre: cierre.empleado.contact_name,
      dni: cierre.empleado.cf_dni_cliente,
      inicioTurno: cierre.empleado.horaEntrada,
      finTurno: salida
    };

    let dataReport = {
      pax: pax,
      desvio: desvio,
      detalleCierre: fondo.detalleCierre,
      empleado: empleado,
      desglose: moneda,
      totalFondo: totalFondo,
      detalleAuto: detalleAuto,
      cierre: cierre
    };

    resolve(dataReport);
  });
};

const generateExelSheetOne = (calculateTotal, res) => {
  return new Promise((resolve, reject) => {
    let fileName = `cierreCaja${date}.xlsx`;
    let workbook = new Excel.Workbook();
    workbook.xlsx.readFile(TEMPLATE).then(() => {
      const worksheet = workbook.getWorksheet("Ar1");

      let row;

      let detalleAuto = calculateTotal.detalleAuto;
      let detalleCierre = calculateTotal.detalleCierre;
      let apertura = calculateTotal.totalFondo;
      let entradasEfectivo = 0; //POR DEFINIR CUANDO CAJA TENGA ENTRADAS
      let detalleArqMoneda = calculateTotal.desglose;
      let salidas = detalleArqMoneda.vales;

      row = worksheet.getRow(7); //ROWS 1...N
      row.getCell(D).value = detalleArqMoneda.dosCientosSolesBillete;
      row.commit();
      row.getCell(E).value = detalleArqMoneda.dosCientosSolesBillete * 200;
      row.commit();

      row = worksheet.getRow(8); //ROWS 1...N
      row.getCell(D).value = detalleArqMoneda.cienSolesBillete;
      row.commit();
      row.getCell(E).value = parseFloat(
        (detalleArqMoneda.cienSolesBillete * 100).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(9); //ROWS 1...N
      row.getCell(D).value = detalleArqMoneda.cincuentaSolesBillete;
      row.commit();
      row.getCell(E).value = parseFloat(
        (detalleArqMoneda.cincuentaSolesBillete * 50).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(10); //ROWS 1...N
      row.getCell(D).value = detalleArqMoneda.veinteSolesBillete;
      row.commit();
      row.getCell(E).value = parseFloat(
        (detalleArqMoneda.veinteSolesBillete * 20).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(11); //ROWS 1...N
      row.getCell(D).value = detalleArqMoneda.diezSolesBillete;
      row.commit();
      row.getCell(E).value = parseFloat(
        (detalleArqMoneda.diezSolesBillete * 10).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(12); //ROWS 1...N
      row.getCell(D).value = detalleArqMoneda.cincoSolesMoneda;
      row.commit();
      row.getCell(E).value = parseFloat(
        (detalleArqMoneda.cincoSolesMoneda * 5).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(13); //ROWS 1...N
      row.getCell(D).value = detalleArqMoneda.dosSolesMoneda;
      row.commit();
      row.getCell(E).value = parseFloat(
        (detalleArqMoneda.dosSolesMoneda * 2).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(14); //ROWS 1...N
      row.getCell(D).value = detalleArqMoneda.unSolMoneda;
      row.commit();
      row.getCell(E).value = detalleArqMoneda.unSolMoneda;
      row.commit();

      row = worksheet.getRow(15); //ROWS 1...N
      row.getCell(D).value = detalleArqMoneda.cincuentaCentimosMoneda;
      row.commit();
      row.getCell(E).value = parseFloat(
        (detalleArqMoneda.cincuentaCentimosMoneda * 0.5).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(16); //ROWS 1...N
      row.getCell(D).value = detalleArqMoneda.veinteCentimosMoneda;
      row.commit();
      row.getCell(E).value = parseFloat(
        (detalleArqMoneda.veinteCentimosMoneda * 0.2).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(17); //ROWS 1...N
      row.getCell(D).value = detalleArqMoneda.diezCentimosMoneda;
      row.commit();
      row.getCell(E).value = parseFloat(
        (detalleArqMoneda.diezCentimosMoneda * 0.1).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(18); //TOTAL PIEZAS MONETARIAS
      row.getCell(E).value = parseFloat(
        detalleCierre.totalEfectivoLocalArq.toFixed(2)
      );
      row.commit();

      //----------------------------------CUADRO 1 GARANTIZADO-----------------------------

      row = worksheet.getRow(32); //TOTAL EFECTIVO MONEDA LOCAL
      row.getCell(D).value = (
        parseFloat(detalleCierre.totalEfectivoLocal) - apertura
      ).toFixed(2);
      row.commit();

      row = worksheet.getRow(33); //TOTAL SALIDAS DE EFECTIVO
      row.getCell(D).value = entradasEfectivo;
      row.commit();

      row = worksheet.getRow(34); //TOTAL SALIDAS DE EFECTIVO
      row.getCell(D).value = salidas;
      row.commit();

      row = worksheet.getRow(35); //FONDO DE APERTURA
      row.getCell(D).value = apertura;
      row.commit();

      row = worksheet.getRow(36); //FONDO DE APERTURA
      row.getCell(D).value = parseFloat(
        (
          parseFloat(detalleCierre.totalEfectivoLocal) +
          entradasEfectivo -
          salidas
        ).toFixed(2)
      );
      row.commit();

      //----------------------------------CUADRO 3 y 5 GARANTIZADO-----------------------------

      let totalPlanilla = 0;
      let totalCanje = 0;
      let totalInterno = 0;
      let totalInv = 0;

      let totalPlanillaArq = 0;
      let totalCanjeArq = 0;
      let totalInternoArq = 0;
      let totalInvArq = 0;

      let totalPlanillaOper = 0;
      let totalCanjeOper = 0;
      let totalInternoOper = 0;
      let totalInvOper = 0;

      let orderDetails = []; //PARA ARTICULO TOP VENTAS

      detalleAuto.forEach(pago => {
        orderDetails.push(pago.orderDetails);
        if (pago.tipoPago.toLowerCase().includes("planilla")) {
          totalPlanilla = totalPlanilla + pago.monto;
          totalPlanillaOper = totalPlanillaOper + 1;
        } else if (pago.tipoPago.toLowerCase().includes("canje")) {
          totalCanje = totalCanje + pago.monto;
          totalCanjeOper = totalCanjeOper + 1;
        } else if (pago.tipoPago.toLowerCase().includes("consumo")) {
          totalInterno = totalInterno + pago.monto;
          totalInternoOper = totalInternoOper + 1;
        } else if (pago.tipoPago.toLowerCase().includes("invitacion")) {
          totalInv = totalInv + pago.monto;
          totalInvOper = totalInvOper + 1;
        }
      });

      orderDetails = _.flatten(orderDetails);

      let arqOtros = calculateTotal.detalleCierre.arqOtros;

      arqOtros.forEach(tpago => {
        if (tpago.tipoPago.toLowerCase().includes("planilla")) {
          totalPlanillaArq = totalPlanillaArq + tpago.monto;
        } else if (tpago.tipoPago.toLowerCase().includes("canje")) {
          totalCanjeArq = totalCanjeArq + tpago.monto;
        } else if (tpago.tipoPago.toLowerCase().includes("consumo")) {
          totalInternoArq = totalInternoArq + tpago.monto;
        } else if (tpago.tipoPago.toLowerCase().includes("invitacion")) {
          totalInvArq = totalInvArq + tpago.monto;
        }
      });

      row = worksheet.getRow(21); //TOTAL EFECTIVO MONEDA LOCAL
      row.getCell(D).value = (
        parseFloat(detalleCierre.totalEfectivoLocal) - apertura
      ).toFixed(2);
      row.commit();
      row = worksheet.getRow(35); //TOTAL EFECTIVO MONEDA LOCAL
      row.getCell(K).value = parseFloat(detalleCierre.totalEfectivoLocal);
      row.commit();
      row = worksheet.getRow(35); //TOTAL EFECTIVO MONEDA LOCAL DECLARADO
      row.getCell(J).value = parseFloat(
        detalleCierre.totalEfectivoLocalArq.toFixed(2)
      );
      row.commit();
      row = worksheet.getRow(21); //TOTAL OPERACIONES EFECTIVO MONEDA LOCAL
      row.getCell(E).value = detalleCierre.operEfectivoLocal;
      row.commit();
      row = worksheet.getRow(21); //TOTAL OPERACIONES EFECTIVO MONEDA LOCAL DECLARADO
      row.getCell(L).value = detalleCierre.totalEfectivoLocalArq;
      row.commit();
      row = worksheet.getRow(35); //TOTAL OPERACIONES EFECTIVO MONEDA LOCAL DECLARADO
      row.getCell(L).value = parseFloat(
        (
          detalleCierre.totalEfectivoLocalArq -
          parseFloat(detalleCierre.totalEfectivoLocal)
        ).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(22); //TOTAL EFECTIVO MONEDA EXTRANJERA
      row.getCell(D).value = detalleCierre.totalEfectivoDolar;
      row.commit();
      row = worksheet.getRow(36); //TOTAL EFECTIVO MONEDA EXTRANJERA
      row.getCell(K).value = detalleCierre.totalEfectivoDolar;
      row.commit();
      row = worksheet.getRow(36); //TOTAL EFECTIVO MONEDA EXTRANJERA
      row.getCell(J).value = 0; //ARQUEO DOLAR
      row.commit();
      row = worksheet.getRow(22); //TOTAL OPERACIONES EFECTIVO MONEDA EXTRANJERA
      row.getCell(E).value = 0; //DETERMINAR operEfectivoLocal
      row.commit();
      row = worksheet.getRow(22); //TOTAL OPERACIONES EFECTIVO MONEDA EXTRANJERA DECLARADO
      row.getCell(L).value = 0;
      row.commit();
      row = worksheet.getRow(36); //TOTAL OPERACIONES EFECTIVO MONEDA EXTRANJERA DECLARADO
      row.getCell(L).value = 0; //CALCULAR
      row.commit();

      row = worksheet.getRow(23); //TOTAL VISA
      row.getCell(D).value = parseFloat(detalleCierre.totalPosVisa);
      row.commit();
      row = worksheet.getRow(37); //TOTAL VISA
      row.getCell(K).value = parseFloat(detalleCierre.totalPosVisa);
      row.commit();
      row = worksheet.getRow(37); //TOTAL VISA
      row.getCell(J).value = parseFloat(detalleCierre.totalPosVisaArq);
      row.commit();
      row = worksheet.getRow(23); //TOTAL OPERACIONES VISA
      row.getCell(E).value = detalleCierre.operPosVisa;
      row.commit();
      row = worksheet.getRow(23); //TOTAL VISA
      row.getCell(L).value = parseFloat(
        detalleCierre.totalPosVisaArq.toFixed(2)
      );
      row.commit();
      row = worksheet.getRow(37); //TOTAL VISA
      row.getCell(L).value =
        parseFloat(detalleCierre.totalPosVisaArq.toFixed(2)) -
        parseFloat(detalleCierre.totalPosVisa);
      row.commit();

      row = worksheet.getRow(24); //TOTAL MASTER
      row.getCell(D).value = parseFloat(detalleCierre.totalPosMaster);
      row.commit();
      row = worksheet.getRow(38); //TOTAL MASTER
      row.getCell(K).value = parseFloat(detalleCierre.totalPosMaster);
      row.commit();
      row = worksheet.getRow(38); //TOTAL MASTER
      row.getCell(J).value = detalleCierre.totalPosMasterArq;
      row.commit();
      row = worksheet.getRow(24); //TOTAL OPERACIONES MASTER
      row.getCell(E).value = detalleCierre.operPosMaster;
      row.commit();
      row = worksheet.getRow(24); //TOTAL MASTER
      row.getCell(L).value = detalleCierre.totalPosMasterArq;
      row.commit();
      row = worksheet.getRow(38); //TOTAL MASTER
      row.getCell(L).value = parseFloat(
        (
          detalleCierre.totalPosMasterArq -
          parseFloat(detalleCierre.totalPosMaster)
        ).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(25); //TOTAL PLANILLA
      row.getCell(D).value = totalPlanilla;
      row.commit();
      row = worksheet.getRow(39); //TOTAL PLANILLA
      row.getCell(K).value = totalPlanilla;
      row.commit();
      row = worksheet.getRow(39); //TOTAL PLANILLA
      row.getCell(J).value = totalPlanillaArq;
      row.commit();
      row = worksheet.getRow(25); //TOTAL OPERACIONES PLANILLA
      row.getCell(E).value = totalPlanillaOper;
      row.commit();
      row = worksheet.getRow(25); //TOTAL PLANILLA
      row.getCell(L).value = totalPlanillaArq;
      row.commit();
      row = worksheet.getRow(39); //TOTAL PLANILLA
      row.getCell(L).value = parseFloat(
        (totalPlanillaArq - totalPlanilla).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(26); //TOTAL CANJE
      row.getCell(D).value = totalCanje;
      row.commit();
      row = worksheet.getRow(40); //TOTAL CANJE
      row.getCell(K).value = totalCanje;
      row.commit();
      row = worksheet.getRow(40); //TOTAL CANJE
      row.getCell(J).value = totalCanjeArq;
      row.commit();
      row = worksheet.getRow(26); //TOTAL OPERACIONES CANJE
      row.getCell(E).value = totalCanjeOper;
      row.commit();
      row = worksheet.getRow(26); //TOTAL CANJE
      row.getCell(L).value = totalCanjeArq;
      row.commit();
      row = worksheet.getRow(40); //TOTAL CANJE
      row.getCell(L).value = parseFloat(
        (totalCanjeArq - totalCanje).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(27); //TOTAL CONSUMO INTERNO
      row.getCell(D).value = totalInterno;
      row.commit();
      row = worksheet.getRow(41); //TOTAL CONSUMO INTERNO
      row.getCell(K).value = totalInterno;
      row.commit();
      row = worksheet.getRow(41); //TOTAL CONSUMO INTERNO
      row.getCell(J).value = totalInternoArq;
      row.commit();
      row = worksheet.getRow(27); //TOTAL OPERACIONES CONSUMO INTERNO
      row.getCell(E).value = totalInternoOper;
      row.commit();
      row = worksheet.getRow(27); //TOTAL CONSUMO INTERNO
      row.getCell(L).value = totalInternoArq;
      row.commit();
      row = worksheet.getRow(41); //TOTAL CONSUMO INTERNO
      row.getCell(L).value = parseFloat(
        (totalInternoArq - totalInterno).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(28); //TOTAL INVITACION
      row.getCell(D).value = totalInv;
      row.commit();
      row = worksheet.getRow(42); //TOTAL INVITACION
      row.getCell(K).value = totalInv;
      row.commit();
      row = worksheet.getRow(42); //TOTAL INVITACION
      row.getCell(J).value = totalInvArq;
      row.commit();
      row = worksheet.getRow(28); //TOTAL OPERACIONES INVITACION
      row.getCell(E).value = totalInvOper;
      row.commit();
      row = worksheet.getRow(28); //TOTAL INVITACION
      row.getCell(L).value = totalInvArq;
      row.commit();
      row = worksheet.getRow(42); //TOTAL INVITACION
      row.getCell(L).value = parseFloat((totalInvArq - totalInv).toFixed(2));
      row.commit();

      row = worksheet.getRow(29); //TOTAL GENERAL
      row.getCell(D).value = parseFloat(
        (
          detalleCierre.totalEfectivoLocal -
          apertura +
          detalleCierre.totalEfectivoDolar +
          detalleCierre.totalPosVisa +
          detalleCierre.totalPosMaster +
          totalPlanilla +
          totalCanje +
          totalInterno +
          totalInv
        ).toFixed(2)
      );
      row.commit();

      row = worksheet.getRow(29); //TOTAL GENERAL
      row.getCell(L).value = parseFloat(
        (
          detalleCierre.totalEfectivoLocalArq +
          detalleCierre.totalEfectivoDolarArq +
          detalleCierre.totalPosVisaArq +
          detalleCierre.totalPosMasterArq +
          totalPlanillaArq +
          totalCanjeArq +
          totalInternoArq +
          totalInvArq
        ).toFixed(2)
      );
      row.commit();

      let tcalc = parseFloat(
        (
          detalleCierre.totalEfectivoLocal -
          apertura +
          detalleCierre.totalEfectivoDolar +
          detalleCierre.totalPosVisa +
          detalleCierre.totalPosMaster +
          totalPlanilla +
          totalCanje +
          totalInterno +
          totalInv
        ).toFixed(2)
      );
      row = worksheet.getRow(43); //TOTAL GENERAL
      row.getCell(K).value = parseFloat(
        (
          detalleCierre.totalEfectivoLocal +
          detalleCierre.totalEfectivoDolar +
          detalleCierre.totalPosVisa +
          detalleCierre.totalPosMaster +
          totalPlanilla +
          totalCanje +
          totalInterno +
          totalInv
        ).toFixed(2)
      );
      row.commit();

      let tarc = parseFloat(
        (
          detalleCierre.totalEfectivoLocalArq +
          detalleCierre.totalEfectivoDolarArq +
          detalleCierre.totalPosVisaArq +
          detalleCierre.totalPosMasterArq +
          totalPlanillaArq +
          totalCanjeArq +
          totalInternoArq +
          totalInvArq
        ).toFixed(2)
      );

      let tcalcResumen =
        detalleCierre.totalEfectivoLocalArq +
        0 +
        detalleCierre.totalPosVisaArq +
        detalleCierre.totalPosMasterArq +
        totalPlanillaArq +
        totalCanjeArq +
        totalInternoArq +
        totalInvArq;

      row = worksheet.getRow(43); //TOTAL GENERAL
      row.getCell(J).value = tcalcResumen;

      row.commit();

      let tventas = (
        detalleCierre.operEfectivoLocal +
        0 +
        detalleCierre.operPosVisa +
        detalleCierre.operPosMaster +
        totalPlanillaOper +
        totalCanjeOper +
        totalInternoOper +
        totalInvOper
      ).toFixed(2);

      row = worksheet.getRow(29); //TOTAL OPERACIONES GENERAL FALTA operEfectivoDolar
      row.getCell(E).value = tventas;

      row.commit();

      row = worksheet.getRow(43); //CUADRE DE EFECTIVO
      row.getCell(L).value = parseFloat(
        (
          worksheet.getCell("J43").value - worksheet.getCell("K43").value
        ).toFixed(2)
      );
      row.commit();

      //----------------------------------CUADRO 2 GARANTIZADO-----------------------------

      row = worksheet.getRow(55); //CAJERO
      row.getCell(C).value = `Cajero: ${calculateTotal.empleado.nombre}`;
      row.commit();
      row = worksheet.getRow(56); //DNI
      row.getCell(C).value = `Dni: ${calculateTotal.empleado.dni}`;
      row.commit();
      row = worksheet.getRow(3); //INICIO DE TURNO
      row.getCell(D).value = `${calculateTotal.empleado.finTurno}`;
      row.commit();
      row = worksheet.getRow(4); //FIN DE TURNO
      row.getCell(D).value = `${calculateTotal.empleado.inicioTurno}`;
      row.commit();

      //----------------------------------EMPLEADO GARANTIZADO-----------------------------

      let entradas = [];
      let fondos = [];
      let adicionales = [];
      let bebidasCal = [];
      let bebidasFri = [];
      let postres = [];
      let desayunos = [];

      orderDetails.forEach(articulo => {
        let categoria = articulo.categoria.toLowerCase();
        if (
          categoria.includes("entradas") ||
          categoria.includes("menu (entradas)") ||
          categoria.includes("ensalada") ||
          categoria.includes("piqueo") ||
          categoria.includes("tapa") ||
          categoria.includes("sopa")
        ) {
          entradas.push(articulo);
        } else if (
          categoria.includes("fondo") ||
          categoria.includes("menu (fondos)") ||
          categoria.includes("menu del dia") ||
          categoria.includes("planchas") ||
          categoria.includes("pastas")
        ) {
          fondos.push(articulo);
        } else if (categoria.includes("adicionales")) {
          adicionales.push(articulo);
        } else if (
          categoria.includes("bebidas frias") ||
          categoria.includes("cervezas") ||
          categoria.includes("vinos") ||
          categoria.includes("menu (bebidas)") ||
          categoria.includes("cocteleria")
        ) {
          bebidasFri.push(articulo);
        } else if (categoria.includes("bebidas calientes")) {
          bebidasCal.push(articulo);
        } else if (
          categoria.includes("postres carta") ||
          categoria.includes("postres barra")
        ) {
          postres.push(articulo);
        } else if (
          categoria.includes("desayunos (carta)") ||
          categoria.includes("desayunos (barra)") ||
          categoria.includes("combos")
        ) {
          desayunos.push(articulo);
        }
      });

      function mode(arr) {
        return arr
          .sort((a, b) => {
            return (
              arr.filter(v => v.sku === a.sku).length -
              arr.filter(v => v.sku === b.sku).length
            );
          })
          .pop();
      }

      let entrada = entradas.length > 0 ? mode(entradas) : "";
      let fondo = fondos.length > 0 ? mode(fondos) : "";
      let adicional = adicionales.length > 0 ? mode(adicionales) : "";
      let bebidaFri = bebidasFri.length > 0 ? mode(bebidasFri) : "";
      let bebidaCal = bebidasCal.length > 0 ? mode(bebidasCal) : "";
      let postre = postres.length > 0 ? mode(postres) : "";
      let desayuno = desayunos.length > 0 ? mode(desayunos) : "";

      row = worksheet.getRow(8); //entrada
      row.getCell(J).value = entrada.sku;
      row.commit();

      row = worksheet.getRow(9); //fondo
      row.getCell(J).value = fondo.sku;
      row.commit();

      row = worksheet.getRow(10); //Fadicional
      row.getCell(J).value = adicional.sku;
      row.commit();

      row = worksheet.getRow(11); //bebidaFri
      row.getCell(J).value = bebidaFri.sku;
      row.commit();

      row = worksheet.getRow(12); //bebidaCal
      row.getCell(J).value = bebidaCal.sku;
      row.commit();

      row = worksheet.getRow(13); //postre
      row.getCell(J).value = postre.sku;
      row.commit();

      row = worksheet.getRow(14); //desayuno
      row.getCell(J).value = desayuno.sku;
      row.commit();

      //--------------------------- TOP VENTAS GARANTIZADO------------------------------

      row = worksheet.getRow(39); //VENTAS DELIVERY NO HAY
      row.getCell(E).value = 0;
      row.commit();
      row = worksheet.getRow(40); //VENTAS DELIVERY NO HAY
      row.getCell(E).value = 0;
      row.commit();
      row = worksheet.getRow(41); //VENTAS DELIVERY NO HAY
      row.getCell(E).value = 0;
      row.commit();
      row = worksheet.getRow(42); //VENTAS DELIVERY NO HAY
      row.getCell(E).value = 0;
      row.commit();
      row = worksheet.getRow(43); //VENTAS DELIVERY NO HAY
      row.getCell(E).value = 0;
      row.commit();

      row = worksheet.getRow(39); //VENTAS DELIVERY NO HAY
      row.getCell(F).value = worksheet.getCell("D29").value;
      row.commit();
      row = worksheet.getRow(40); //VENTAS DELIVERY NO HAY
      row.getCell(F).value = detalleAuto.length - 1;
      row.commit();
      row = worksheet.getRow(41); //VENTAS DELIVERY NO HAY
      row.getCell(F).value = calculateTotal.pax;
      row.commit();
      row = worksheet.getRow(42); //VENTAS DELIVERY NO HAY
      row.getCell(F).value = (
        worksheet.getCell("F39").value / worksheet.getCell("F40").value
      ).toFixed(2);
      row.commit();
      row = worksheet.getRow(43); //VENTAS DELIVERY NO HAY
      row.getCell(F).value = (
        worksheet.getCell("F39").value / worksheet.getCell("F41").value
      ).toFixed(2);

      row.commit();

      row = worksheet.getRow(39); //VENTAS DELIVERY NO HAY
      row.getCell(G).value =
        worksheet.getCell("D39").value +
        worksheet.getCell("E39").value +
        worksheet.getCell("F39").value;
      row.commit();
      row = worksheet.getRow(40); //VENTAS DELIVERY NO HAY
      row.getCell(G).value = detalleAuto.length - 1;
      row.commit();
      row = worksheet.getRow(41); //VENTAS DELIVERY NO HAY
      row.getCell(G).value = calculateTotal.pax;
      row.commit();
      row = worksheet.getRow(42); //VENTAS DELIVERY NO HAY
      row.getCell(G).value = (
        worksheet.getCell("F39").value / worksheet.getCell("F40").value
      ).toFixed(2);

      row.commit();
      row = worksheet.getRow(43); //VENTAS DELIVERY NO HAY
      row.getCell(G).value = (
        worksheet.getCell("F39").value / worksheet.getCell("F41").value
      ).toFixed(2);
      row.commit();

      workbook.xlsx.writeFile(`./reports/${fileName}`).then(async () => {
        const thePdf = await getPdf([fileName], res).catch(err =>
          console.log(err)
        );
        if (thePdf) resolve("done");
      });
      // resolve("done");
    });
  });
};

const getPdf = (fileNames, res) => {
  let data = {
    fileNames: fileNames,
    cierre: cierre
  };
  return new Promise((resolve, reject) => {
    axios
      .post(`http://${SUB_END_POINT}:3001/api/getpdf`, data, {
        headers: { "Access-Control-Allow-Origin": "*" },
        responseType: "json"
      })
      .then(response => {
        const myPDF = `${response.data}.pdf`;
        res.setHeader(
          "Content-disposition",
          'attachment; filename="' + myPDF + '"'
        );
        res.setHeader("Content-type", "blob");
        res.download(myPDF);
        resolve("done");
      })
      .catch(error => {
        reject(error.message);
      });
  });
};

const setUpArqueo = async (res, cierres) => {
  const cierre = await getCierre(cierres).catch(err => console.log(err));
  if (cierre) {
    const calculateTotal = await calculateTotals(cierre).catch(err =>
      console.log(err)
    );
    if (calculateTotal) {
      const printToExcel = await generateExelSheetOne(
        calculateTotal,
        res
      ).catch(err => console.log(err));
      if (printToExcel) {
        console.log("done");
      }
    }
  }
};

router.get("/api/getCierre/:cierre", (req, res) => {
  cierre = req.params.cierre;
  setUpArqueo(res, cierre);
});

module.exports = router;
