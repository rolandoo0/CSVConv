from cmath import nan
import csv
import os
import sys
import pandas as pd
import datetime
from pandas import ExcelWriter
from openpyxl import load_workbook
import time
import colorama
from colorama import Fore, Back, Style
import codecs

columnNames = [
    "order-id",
    "buyer-phone-number",
    "product-name",
    "quantity-to-ship",
    "recipient-name",
    "ship-address-1",
    "ship-address-2",
    "ship-city",
    "ship-state",
    "ship-postal-code",
]


def isALPHA(string):
    status = string.isalpha()
    return status


def formatNAMES(row):
    try:
        splitName = row.split(" ")
    except:
        return
    finalName = {}
    for i in splitName:
        if not isALPHA(i):
            try:
                newRow = i.replace("Ã", "i")
                if isALPHA(newRow):
                    if newRow not in finalName.values():
                        finalName[splitName.index(i)] = newRow
            except:
                pass
            try:
                newRow = i.replace("Ô", "o")
                if isALPHA(newRow):
                    if newRow not in finalName.values():
                        finalName[splitName.index(i)] = newRow
            except:
                pass
            try:
                if isALPHA(newRow):
                    if newRow not in finalName.values():
                        finalName[splitName.index(i)] = newRow
                else:
                    newRow = i.replace("Ã¡", "a")
            except:
                pass
            try:
                if isALPHA(newRow):
                    if newRow not in finalName.values():
                        finalName[splitName.index(i)] = newRow
                else:
                    newRow = i.replace("Ã©", "e")
            except:
                pass
            try:
                if isALPHA(newRow):
                    if newRow not in finalName.values():
                        finalName[splitName.index(i)] = newRow
                else:
                    newRow = i.replace("Ã±", "n")
            except:
                pass
            try:
                if isALPHA(newRow):
                    if newRow not in finalName.values():
                        finalName[splitName.index(i)] = newRow
                else:
                    newRow = i.replace("Ã³", "o")
            except:
                pass
            try:
                if isALPHA(newRow):
                    if newRow not in finalName.values():
                        finalName[splitName.index(i)] = newRow
                else:
                    newRow = i.replace("Ãº", "u")
            except:
                pass
        elif splitName.index(i) == len(splitName) - 1:
            for index in finalName:
                splitName[index] = finalName[index]
    return " ".join(splitName)


headers = {
    "Observaciones 2": " ",
    "Observaciones 3": " ",
    "Observaciones 4": " ",
    "Retorno": "N",
    "8:30": "N",
    "Sábado": "N",
    "Gestión/Tramite": "N",
    "Ok 15": "N",
    "Prepagado": "N",
    "Tipo Seguro": "N",
    "Imp. Seg.": "0",
    "Tipo Ealerta": " ",
    "Ealerta": " ",
    "Alto": " ",
    "Ancho": " ",
    "Largo": " ",
    "Contenido": " ",
    "Valor Declarado": " ",
    "Digitalizar/Almacenar": "N",
    "Alb.Cli.": "0",
    "Ins. Adi.": " ",
    "Ins1": " ",
    "Ins2": " ",
    "Ins3": " ",
    "Ins4": " ",
    "Ins5": " ",
    "Ins6": " ",
    "Ins7": " ",
    "Ins8": " ",
    "Ins9": " ",
    "Ins10": " ",
    "Ins11": " ",
    "Ins12": " ",
    "Ins13": " ",
    "Ins14": " ",
    "Ins15": " ",
    "Tipo PreAlerta": " ",
    "Modo Alerta": " ",
    "PreAlerta": " ",
    "PreAlerta Mensaje": " ",
    "Tipo PreAlerta1": " ",
    "Modo Alerta1": " ",
    "PreAlerta1": " ",
    "PreAlerta Mensaje1": " ",
    "Tipo PreAlerta2": " ",
    "Modo Alerta2": " ",
    "PreAlerta2": " ",
    "PreAlerta Mensaje2": " ",
    "Tipo PreAlerta3": " ",
    "Modo Alerta3": " ",
    "PreAlerta3": " ",
    "PreAlerta Mensaje3": " ",
    "Tipo PreAlerta4": " ",
    "Modo Alerta4": " ",
    "PreAlerta4": " ",
    "PreAlerta Mensaje4": " ",
    "Horario concertado": " ",
    "incio horario concertado": " ",
    "fin horario concertado": " ",
    "referencias multiples": " ",
}


def convertToCSV(input_file):
    masterRows = []
    with open(input_file, "rb") as csv_in_file:
        for ln in csv_in_file:
            decoded = False
            line = ""
            for cp in ("cp1252", "cp850", "utf-8", "utf8"):
                try:
                    line = ln.decode(cp)
                    decoded = True
                    break
                except UnicodeDecodeError:
                    pass
            if decoded:
                try:
                    line = line.replace("0xae", "")
                except:
                    pass
                try:
                    line = line.replace("├", "")
                except:
                    pass
                try:
                    line = line.replace("n┬║", "")
                except:
                    pass
                try:
                    line = line.replace("N┬║52", "")
                except:
                    pass
                try:
                    line = line.replace("1┬║D", "")
                except:
                    pass
                try:
                    line = line.replace("▒", "")
                except:
                    pass
                for i in line:
                    if (
                        not i.isalpha()
                        and not i.isnumeric()
                        and i != "-"
                        and i != "\t"
                        and i != " "
                        and i != ":"
                        and i != "+"
                        and i != "@"
                        and i != "."
                    ):
                        line = line.replace(i, "")
                masterRows.append(line)
        with open("output.csv", "w", newline="") as csv_out_file:
            filewriter = csv.writer(csv_out_file)
            try:
                for row in masterRows:
                    indexxx = masterRows.index(row)
                    newRow = []
                    length = len(row)
                    if length > 1:
                        string = ""
                        for i in range(length):
                            string = string + row[i]
                        rows = string.split("\t")
                        for subRow in rows:
                            newRow.append(subRow)
                    else:
                        rows = row[0].split("\t")
                        for subRow in rows:
                            newRow.append(subRow)
                    filewriter.writerow(newRow)
            except Exception as e:
                print(e)
                pass
    return rows


def readCSV(outputFile, inputFile):
    columnsToDelete = []
    try:
        df = pd.read_csv(outputFile)
        dictionary = df.to_dict()
    except:
        """question = input(Fore.YELLOW + "Se detectaron errores en el archivo, quisiera obtener el mayor número de filas posibles?\nDe lo contrario se obtendran todas las filas pero con otro formato. (y/n) " + Fore.RESET)
        if question == "y":
            convertAFTERQUESTION(inputFile)
            df = pd.read_csv(outputFile)
        else:"""
        df = pd.read_csv(outputFile, encoding="unicode_escape")
        dictionary = df.to_dict()
    for column in df.columns:
        if column not in columnNames:
            df.pop(column)
    return dictionary


def removeUnusedColumns(dictionary):
    columnsToRemove = []
    for column in dictionary:
        if column not in columnNames:
            columnsToRemove.append(column)
    for column in columnsToRemove:
        dictionary.pop(column)
    return dictionary


def addNewColumns(dictionary):
    """with ExcelWriter("outputExcel.xlsx") as writer:
    df.to_excel(writer, sheet_name="Sheet1")
    workbook: ExcelWriter = writer.book
    worksheet: ExcelWriter = writer.sheets["Sheet1"]"""
    removeUnusedColumns(dictionary)
    dictionary["Nom. Entrega"] = dictionary.pop("recipient-name")
    dictionary["Dirección Ent."] = dictionary.pop("ship-address-1")
    dictionary["CPEntrega"] = dictionary.pop("ship-postal-code")
    dictionary["Población Ent."] = dictionary.pop("ship-city")
    dictionary["Provincia Ent."] = dictionary.pop("ship-state")
    dictionary["Teléfono Ent."] = dictionary.pop("buyer-phone-number")
    dictionary["Observaciones 1"] = dictionary.pop("product-name")
    lengthOfRows = len(dictionary["Nom. Entrega"])
    numbers = {}
    for i in range(0, lengthOfRows):
        numbers[i] = "37"
    dictionary["Nº Abonado"] = numbers

    departamentos = {}
    for i in range(0, lengthOfRows):
        departamentos[i] = " "
    dictionary["Departamento"] = departamentos

    servicios = {}
    for i in range(0, lengthOfRows):
        servicios[i] = "5"
    dictionary["Servicio"] = servicios

    tipoCobro = {}
    for i in range(0, lengthOfRows):
        tipoCobro[i] = "0"
    dictionary["Tipo Cobro"] = tipoCobro

    excesos = {}
    for i in range(0, lengthOfRows):
        excesos[i] = " "
    dictionary["Excesos"] = excesos
    dictionary["Referencia"] = dictionary.pop("order-id")
    bagsAndPacks = {}
    for i in range(0, lengthOfRows):
        bagsAndPacks[i] = "1"
    dictionary["Bag/Paq"] = bagsAndPacks

    bultos = {}
    for i in range(0, lengthOfRows):
        bultos[i] = "1"
    dictionary["Bultos"] = bultos

    kilos = {}
    for i in range(0, lengthOfRows):
        kilos[i] = "1"
    dictionary["Kilos"] = kilos

    nomEntrega = dictionary.pop("Nom. Entrega")
    dictionary["Nom. Entrega"] = nomEntrega

    departamentoEnt = {}
    for i in range(0, lengthOfRows):
        departamentoEnt[i] = " "
    dictionary["Departamento Ent."] = departamentoEnt

    personaEnt = {}
    for i in range(0, lengthOfRows):
        personaEnt[i] = " "
    dictionary["Persona Ent."] = personaEnt

    direccionEnt = dictionary.pop("Dirección Ent.")
    dictionary["Dirección Ent."] = direccionEnt

    paisEnt = {}
    for i in range(0, lengthOfRows):
        paisEnt[i] = "ES"
    dictionary["Pais Ent."] = paisEnt

    codigoPostalEnt = dictionary.pop("CPEntrega")
    dictionary["CPEntrega"] = codigoPostalEnt

    poblaEnt = dictionary.pop("Población Ent.")
    dictionary["Población Ent."] = poblaEnt

    provinciaEnt = dictionary.pop("Provincia Ent.")
    dictionary["Provincia Ent."] = provinciaEnt

    telefonoEnt = dictionary.pop("Teléfono Ent.")
    dictionary["Teléfono Ent."] = telefonoEnt

    impRee = {}
    for i in range(0, lengthOfRows):
        impRee[i] = "0"
    dictionary["Imp. Ree."] = impRee

    tipoRee = {}
    for i in range(0, lengthOfRows):
        tipoRee[i] = "N"
    dictionary["Tipo Ree."] = tipoRee

    observaciones = dictionary.pop("Observaciones 1")
    dictionary["Observaciones 1"] = observaciones
    quantityDF = dictionary.pop("quantity-to-ship")
    for r in range(21, 82):
        for header in headers:
            columnInfo = {}
            for i in range(0, lengthOfRows):
                columnInfo[i] = headers[header]
            dictionary[header] = columnInfo
            headers.pop(header)
            break
    dictionary["quantity-to-ship"] = quantityDF
    df = pd.DataFrame.from_dict(dictionary)
    df = df.rename(
        columns={
            "Tipo PreAlerta1": "Tipo PreAlerta",
            "Modo Alerta1": "Modo Alerta",
            "PreAlerta1": "PreAlerta",
            "PreAlerta Mensaje1": "PreAlerta Mensaje",
            "Tipo PreAlerta2": "Tipo PreAlerta",
            "Modo Alerta2": "Modo Alerta",
            "PreAlerta2": "PreAlerta",
            "PreAlerta Mensaje2": "PreAlerta Mensaje",
            "Tipo PreAlerta2": "Tipo PreAlerta",
            "Modo Alerta3": "Modo Alerta",
            "PreAlerta3": "PreAlerta",
            "PreAlerta Mensaje3": "PreAlerta Mensaje",
            "Tipo PreAlerta3": "Tipo PreAlerta",
            "Modo Alerta4": "Modo Alerta",
            "PreAlerta4": "PreAlerta",
            "PreAlerta Mensaje4": "PreAlerta Mensaje",
        }
    )
    indexLIST = []
    index = 0
    indexL = list(df.loc[df["quantity-to-ship"] > 1].index)
    for i in indexL:
        if i in indexLIST and indexL.index(i) != len(indexL) - 1:
            pass
        else:
            name = df.loc[i, "Observaciones 1"]
            if (
                "Bag in Box" in name
                and "5L" in name
                and "15L" not in name
                or "5 Litros" in name
            ):
                try:
                    quantity = df.loc[i, "quantity-to-ship"]
                    quantityNumber = quantity.numerator
                except:
                    quantityNumber = int(df.loc[i, "quantity-to-ship"])
                divisibleBy = quantityNumber % 2
                if divisibleBy == 0:
                    indexToAdd = int((quantityNumber / 2) - 1)
                    indexLIST.append(str(int(i) + indexToAdd))
                else:
                    indexLIST.append(str(int(i)))
            elif "Bag in Box 15L" in name:
                indexLIST.append(str(int(i)))
            else:
                indexLIST.append(str(int(i) + index))
                index += 1
    for row in df["quantity-to-ship"]:
        if row == 1:
            pass
        else:
            try:
                rowIndex = indexLIST[0]
            except:
                break
            rowToDuplicate = df.iloc[int(rowIndex)]
            time.sleep(0.2)
            prodNameAtRowIndex = rowToDuplicate.get(
                "Observaciones 1"
            )  # df.iloc[str(rowIndex), "Observaciones 1"]
            if "BAG IN BOX 15L" in prodNameAtRowIndex.upper() and row > 1:
                quantityToGenerate = int(row - 1)
                for i in range(quantityToGenerate):
                    df = insertRow(int(rowIndex) + 1, df, rowToDuplicate)

            elif (
                "BAG IN BOX 5L"
                or "BAG IN BOX"
                and "5 LITROS" in prodNameAtRowIndex.upper()
            ):
                if row <= 2:
                    indexLISTquant = list(
                        df.loc[df["Observaciones 1"] == prodNameAtRowIndex].index
                    )
                    if row == 1:
                        pass
                    if row == 2:
                        newname = prodNameAtRowIndex.upper()
                        if "BLANCO" in newname and "VERDEJO" not in newname:
                            newname = "BAG IN BOX 5L BLANCO JOVEN"
                        elif "ROSADO" in newname:
                            newname = "BAG IN BOX 5L ROSADO JOVEN"
                        elif "PREMIUM" in newname:
                            newname = "BAG IN BOX 5L TINTO PREMIUM"
                        elif "RECOMENDADO" in newname:
                            newname = "BAG IN BOX 5L TINTO RECOMENDADO"
                        elif "VERDEJO" in newname and "BLANCO" not in newname:
                            newname = "BAG IN BOX 5L VERDEJO PAZ VI"
                        elif "NUEVO REINO" in newname:
                            newname = "BAG IN BOX 5L TINTO NUEVO REINO"
                        elif "VERDEJO" in newname and "BLANCO" in newname:
                            newname = "BAG IN BOX 5L BLANCO VERDEJO"
                        df.loc[
                            df["Observaciones 1"] == prodNameAtRowIndex,
                            "Observaciones 1",
                        ] = (
                            "PACK(2) " + newname
                        )
                else:
                    remainder = row % 2
                    index = 0
                    if remainder == 0:
                        division = (row / 2) - 1
                        for i in range(int(division)):
                            df = insertRow(
                                int(rowIndex) + 1 + index, df, rowToDuplicate
                            )
                        listWhereProds = list(
                            df.loc[df["Observaciones 1"] == prodNameAtRowIndex].index
                        )
                        for i in listWhereProds:
                            newname = df.loc[i]["Observaciones 1"].upper()[:40]
                            if "BLANCO" in newname and "VERDEJO" not in newname:
                                newname = "BAG IN BOX 5L BLANCO JOVEN"
                            elif "ROSADO" in newname:
                                newname = "BAG IN BOX 5L ROSADO JOVEN"
                            elif "PREMIUM" in newname:
                                newname = "BAG IN BOX 5L TINTO PREMIUM"
                            elif "RECOMENDADO" in newname:
                                newname = "BAG IN BOX 5L TINTO RECOMENDADO"
                            elif "VERDEJO" in newname and "BLANCO" not in newname:
                                newname = "BAG IN BOX 5L VERDEJO PAZ VI"
                            elif "NUEVO REINO" in newname:
                                newname = "BAG IN BOX 5L TINTO NUEVO REINO"
                            elif "VERDEJO" in newname and "BLANCO" in newname:
                                newname = "BAG IN BOX 5L BLANCO VERDEJO"
                            df.at[df.index[int(i)], "Observaciones 1"] = (
                                "PACK(2) " + newname
                            )
                    else:
                        newLimit = int(((row - 1) / 2) + remainder)
                        for i in range(newLimit):
                            if i < newLimit - 1:
                                df = insertRow(
                                    int(rowIndex) + 1 + index, df, rowToDuplicate
                                )
                                listWhereProds = list(
                                    df.loc[
                                        df["Observaciones 1"] == prodNameAtRowIndex
                                    ].index
                                )
                                newname = prodNameAtRowIndex.upper()[:40]
                                if "BLANCO" in newname and "VERDEJO" not in newname:
                                    newname = "BAG IN BOX 5L BLANCO JOVEN"
                                elif "ROSADO" in newname:
                                    newname = "BAG IN BOX 5L ROSADO JOVEN"
                                elif "PREMIUM" in newname:
                                    newname = "BAG IN BOX 5L TINTO PREMIUM"
                                elif "RECOMENDADO" in newname:
                                    newname = "BAG IN BOX 5L TINTO RECOMENDADO"
                                elif "VERDEJO" in newname and "BLANCO" not in newname:
                                    newname = "BAG IN BOX 5L VERDEJO PAZ VI"
                                elif "NUEVO REINO" in newname:
                                    newname = "BAG IN BOX 5L TINTO NUEVO REINO"
                                elif "VERDEJO" in newname and "BLANCO" in newname:
                                    newname = "BAG IN BOX 5L BLANCO VERDEJO PAZ VI"
                                df.at[
                                    df.index[int(rowIndex) + 1 + index],
                                    "Observaciones 1",
                                ] = (
                                    "PACK(2) " + newname
                                )
                                index += 1
                            else:
                                pass
            elif (
                "BAG IN BOX 3L" in prodNameAtRowIndex.upper()
                or "BAG IN BOX"
                and "3 LITROS" in prodNameAtRowIndex.upper()
            ):
                if row <= 2:
                    indexLISTquant = list(
                        df.loc[df["Observaciones 1"] == prodNameAtRowIndex].index
                    )
                    if row == 1:
                        pass
                    if row == 2:
                        newname = prodNameAtRowIndex.upper()
                        if "RECOMENDADO" in newname:
                            newname = "BAG IN BOX 3L TINTO RECOMENDADO"
                        elif "BAG IN BOX VINO BLANCO VERDEJO 3 LITROS" in newname:
                            newname = "BAG IN BOX 3L VERDEJO PAZ VI"
                        elif "BAG IN BOX 3L VINO TINTO NUEVO REINO" in newname:
                            newname = "BAG IN BOX 3L TINTO NUEVO REINO"
                        elif "BAG IN BOX VINO BLANCO VERDEJO" in newname:
                            newname = "BAG IN BOX 3L BLANCO VERDEJO"
                        df.loc[
                            df["Observaciones 1"] == prodNameAtRowIndex,
                            "Observaciones 1",
                        ] = (
                            "PACK(2) " + newname
                        )
                else:
                    remainder = row % 2
                    index = 0
                    if remainder == 0:
                        division = (row / 2) - 1
                        for i in range(int(division)):
                            df = insertRow(
                                int(rowIndex) + 1 + index, df, rowToDuplicate
                            )
                        listWhereProds = list(
                            df.loc[df["Observaciones 1"] == prodNameAtRowIndex].index
                        )
                        for i in listWhereProds:
                            newname = df.loc[i]["Observaciones 1"].upper()[:40]
                            if "BAG IN BOX VINO BLANCO VERDEJO 3 LITROS" in newname:
                                newname = "BAG IN BOX 3L VERDEJO PAZ VI"
                            elif "BAG IN BOX 3L VINO TINTOS RECOMENDADO" in newname:
                                newname = "BAG IN BOX 3L TINTO RECOMENDADO"
                            elif "BAG IN BOX 3L VINO TINTO NUEVO REINO" in newname:
                                newname = "BAG IN BOX 3L TINTO NUEVO REINO"
                            elif "BAG IN BOX VINO BLANCO VERDEJO" in newname:
                                newname = "BAG IN BOX 3L BLANCO VERDEJO"
                            df.at[df.index[int(i)], "Observaciones 1"] = (
                                "PACK(2) " + newname
                            )
                    else:
                        newLimit = int(((row - 1) / 2) + remainder)
                        for i in range(newLimit):
                            if i < newLimit - 1:
                                df = insertRow(
                                    int(rowIndex) + 1 + index, df, rowToDuplicate
                                )
                                listWhereProds = list(
                                    df.loc[
                                        df["Observaciones 1"] == prodNameAtRowIndex
                                    ].index
                                )
                                newname = prodNameAtRowIndex.upper()[:40]
                                if "BAG IN BOX VINO BLANCO VERDEJO 3 LITROS" in newname:
                                    newname = "BAG IN BOX 3L VERDEJO PAZ VI"
                                elif "BAG IN BOX 3L VINO TINTOS RECOMENDADO" in newname:
                                    newname = "BAG IN BOX 3L TINTO RECOMENDADO"
                                elif "BAG IN BOX 3L VINO TINTO NUEVO REINO" in newname:
                                    newname = "BAG IN BOX 3L TINTO NUEVO REINO"
                                elif "BAG IN BOX VINO BLANCO VERDEJO" in newname:
                                    newname = "BAG IN BOX 3L BLANCO VERDEJO"
                                df.at[
                                    df.index[int(rowIndex) + 1 + index],
                                    "Observaciones 1",
                                ] = (
                                    "PACK(2) " + newname
                                )
                                index += 1
                            else:
                                pass
            indexLIST.pop(indexLIST.index(rowIndex))

    for row in df["Observaciones 1"]:
        try:
            if "PACK - 2" in row:
                index = 0
                indexes = list(df.loc[df["Observaciones 1"] == row].index)
                quantity = list(
                    df.loc[df["Observaciones 1"] == row, "quantity-to-ship"]
                )
                for r in range(len(indexes)):
                    nameAT = df.iloc[int(indexes[r]) + index]["Nom. Entrega"]
                    # ("Name at {i} is {name}".format(i=indexes[r] + index, name=nameAT))
                    # print('Name at 21' + str(df.iloc[int(21)]["Nom. Entrega"]))
                    # print('Name at 22' + str(df.iloc[int(22)]["Nom. Entrega"]))
                    # print('Name at 23' + str(df.iloc[int(23)]["Nom. Entrega"]))
                    amountToDuplicate = (quantity[r] * 2) - 1
                    for i in range(amountToDuplicate):
                        rowToDuplicate = df.iloc[int(indexes[r]) + index]
                        df = insertRow(int(indexes[r]) + 1 + index, df, rowToDuplicate)
                        index += 1
                break
        except:
            pass
    df.pop("quantity-to-ship")

    for row in df["Nom. Entrega"]:
        try:
            newRow = formatNAMES(row)
            df.loc[df["Nom. Entrega"] == row, "Nom. Entrega"] = newRow.upper()
        except:
            pass

    ##CODIGO PARA PAPA:
    for row in df["Observaciones 1"]:
        try:
            if (
                "ColecciÃ³n FAUNA IBERICA Vino Tinto Recomendado - 8 Botellas de 750 ml"
                in row
            ):
                df.loc[
                    df["Observaciones 1"] == row, "Observaciones 1"
                ] = "C-8 BOTELLAS COLECCION FAUNA IBERICA"
            elif (
                "Vino Tinto 00 SIN ALCOHOL - DOS MUNDOS  Caja de 6 botellas x 075 cl"
                in row
            ):
                df.loc[
                    df["Observaciones 1"] == row, "Observaciones 1"
                ] = "C-6 BOTELLAS SIN ALCOHOL DOS MUNDOS"
            elif (
                "Vino Clarete de mesa cosecheroLos Corzos Caja de Botellas 6 x 750 ml - Total: 4500 ml"
                in row
            ):
                df.loc[
                    df["Observaciones 1"] == row, "Observaciones 1"
                ] = "C-6 BOTELLAS CLARETE JOVEN LOS CORZOS"
            """elif "C-8 Blbala" in row:
                df.loc[
                    df["Observaciones 1"] == row, "Observaciones 1"
                ] = "SARMIENTOS DE VID LOS CORZOS
            """
        except:
            pass

    for row in df["Observaciones 1"]:
        try:
            if "Bag in Box 15L Vino Tinto" in row:
                if "Recomendado" in row:
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 15L TINTO RECOMENDADO"
                elif "PREMIUM" in row:
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 15L TINTO PREMIUM"
                elif "Joven" in row:
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 15L TINTO JOVEN"
            elif "PACK - 2" in row:
                df.loc[
                    df["Observaciones 1"] == row, "Observaciones 1"
                ] = "BAG IN BOX 15L TINTO CAJA BARRICA"
            elif (
                "Bag in Box 5L" in row.upper()
                or "BAG IN BOX 5L" in row.upper()
                or "BAG IN BOX"
                and " 5 LITROS" in row.upper()
            ):
                if "PREMIUM" in row:
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 5L TINTO PREMIUM"
                elif "Recomendado" in row:
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 5L TINTO RECOMENDADO"
                elif "cosechero" in row and "blanco" not in row:
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 5L TINTO CAJA BARRICA"
                elif (
                    "BLANCO" in row.upper()
                    and "PACK(2)" not in row
                    and "VERDEJO" not in row.upper()
                ):
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 5L BLANCO JOVEN"
                elif "ROSADO" in row.upper() and "PACK(2)" not in row:
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 5L ROSADO JOVEN"
                elif "JOVEN" in row.upper():
                    if "PACK(2)" in row:
                        df.loc[
                            df["Observaciones 1"] == row, "Observaciones 1"
                        ] = "PACK(2) BAG IN BOX 5L TINTO JOVEN"
                    else:
                        df.loc[
                            df["Observaciones 1"] == row, "Observaciones 1"
                        ] = "BAG IN BOX 5L TINTO JOVEN"
                elif "verdejo" in row.lower() and "blanco" in row.lower():
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 5L VERDEJO PAZ VI"
            elif "Bag in Box verdejo 15 Litros" in row:
                df.loc[
                    df["Observaciones 1"] == row, "Observaciones 1"
                ] = "BAG IN BOX 15L VERDEJO PAZ VI"
            elif "Bag in Box 15L Vino" in row:
                if "cosechero" in row:
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 15L TINTO CAJA BARRICA"
                elif "Blanco" in row:
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 15L BLANCO JOVEN"
                elif "Rosado" in row:
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 15L ROSADO JOVEN"
                elif "Joven" in row:
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 15L TINTO JOVEN"
            elif "BOTELLAS 6" in row.upper() or "6 BOTELLAS" in row.upper():
                if "PAULUS JOVEN" in row.upper():
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS PAULUS JOVEN RIOJA"
                elif "PAULUS CRIANZA" in row.upper():
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS PAULUS CRIANZA RIOJA"
                elif "SOTONOVILLOS CRIANZA" in row.upper():
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS SOTONOVILLOS CRIANZA RIOJA"
                elif "SIDRA NATURAL" in row.upper():
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS SIDRA NATURAL JAREGUI"
                elif "FAUNA IBERICA" in row.upper():
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS COLECCION FAUNA IBERICA"
                elif "PANJUA" in row.upper():
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS SEÑORÍO DE PANJUA VERDEJO"
                elif (
                    "PREMIUM VINO TINTO"
                    and "BOTELLAS 6" in row.upper()
                    or "6 BOTELLAS" in row.upper()
                    and "C-6" not in row.upper()
                ):
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS TINTO PREMIUM LOS CORZOS"
                elif (
                    "CLARETE"
                    and "BOTELLAS 6" in row.upper()
                    or "6 BOTELLAS" in row.upper()
                    and "C-6" not in row.upper()
                ):
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS CLARETE LOS CORZOS"
                elif (
                    "COSECHERO"
                    and "BOTELLAS 6" in row.upper()
                    or "6 BOTELLAS" in row.upper()
                    and "C-6" not in row.upper()
                ):
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS TINTO LOS CORZOS"
                elif (
                    "BLANCO"
                    and "BOTELLAS 6" in row.upper()
                    or "6 BOTELLAS" in row.upper()
                    and "C-6" not in row.upper()
                ):
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS BLANCO LOS CORZOS"
                elif (
                    "ROSADO"
                    and "BOTELLAS 6" in row.upper()
                    or "6 BOTELLAS" in row.upper()
                    and "C-6" not in row.upper()
                ):
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS ROSADO LOS CORZOS"
                elif (
                    "VERDEJO"
                    and "BOTELLAS 6" in row.upper()
                    or "6 BOTELLAS" in row.upper()
                    and "C-6" not in row.upper()
                ):
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS VERDEJO PAZ VI"
                elif (
                    "TINTO RECOMENDADO LOS CORZOS"
                    and "BOTELLAS 6" in row.upper()
                    or "6 BOTELLAS" in row.upper()
                    and "C-6" not in row.upper()
                ):
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS TINTO RECOMENDADO  B/N LOS CORZOS"
                elif (
                    "EL APARATO VINO BLANCO"
                    and "BOTELLAS 6" in row.upper()
                    or "6 BOTELLAS" in row.upper()
                    and "C-6" not in row.upper()
                ):
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS TE ENSEÑO EL APARATO"
                elif (
                    "NUEVO REINO TINTO"
                    and "BOTELLAS 6" in row.upper()
                    or "6 BOTELLAS" in row.upper()
                    and "C-6" not in row.upper()
                ):
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS TINTO NUEVO REINO"
                elif (
                    "RUFUS"
                    and "BOTELLAS 6" in row.upper()
                    or "6 BOTELLAS" in row.upper()
                    and "C-6" not in row.upper()
                ):
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "C-6 BOTELLAS TINTO RUFUS RIBERA DUERO"
            elif (
                "BAG IN BOX 3L" in row.upper()
                or "BAG IN BOX"
                and "3 LITROS" in row.upper()
            ):
                if "COSECHERO" in row.upper():
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 3L TINTO CAJA BARRICA"
                elif "RECOMENDADO" in row.upper():
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 3L TINTO RECOMENDADO"
                elif "VERDEJO" in row.upper() and "BLANCO" in row.upper():
                    df.loc[
                        df["Observaciones 1"] == row, "Observaciones 1"
                    ] = "BAG IN BOX 3L BLANCO VERDEJO AFRUTADO"
        except:
            pass

    with ExcelWriter(
        "{d}.xlsx".format(d=datetime.datetime.now().strftime("%d-%m-%Y"))
    ) as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
        workbook: ExcelWriter = writer.book
        worksheet: ExcelWriter = writer.sheets["Sheet1"]
        for col_cells in worksheet.iter_cols(min_col=15, max_col=15):
            for cell in col_cells:
                while len(str(cell.value)) < 5:
                    cell.number_format = "@"
                    cell.value = "0" + str(cell.value)
                    writer.save()
        writer.save()


def insertRow(row_number, df, row_value):
    start_upper = 0

    end_upper = row_number

    start_lower = row_number

    end_lower = df.shape[0]

    upper_half = [*range(start_upper, end_upper, 1)]

    lower_half = [*range(start_lower, end_lower, 1)]

    lower_half = [x.__add__(1) for x in lower_half]

    index_ = upper_half + lower_half

    df.index = index_

    df.loc[row_number] = row_value

    df = df.sort_index()
    # print(df)
    return df


def formatCSV(dictionary):
    for row in dictionary["ship-postal-code"]:
        newRow = str(dictionary["ship-postal-code"][row])
        while len(str(newRow)) < 5:
            newRow = "0" + str(newRow)
            dictionary["ship-postal-code"][row] = newRow
        dictionary["ship-postal-code"][row] = str(dictionary["ship-postal-code"][row])
    rowIter = 0
    for row in dictionary["buyer-phone-number"]:
        real_row = dictionary["buyer-phone-number"][row]
        try:
            updatedRow = real_row.replace(" ", "")
        except:
            updatedRow = real_row
        newRow = str(updatedRow)
        while len(str(newRow)) > 9:
            newRow = str(newRow.replace(newRow[0], "", 1))
        rowIter = rowIter + 1
        written = True
        try:
            rowIndex = dictionary["ship-postal-code"][row]
        except:
            dictionary["buyer-phone-number"][row - 1] = "no telefono"
            written = False

        if written == True:
            dictionary["buyer-phone-number"][row] = newRow

        else:
            pass

    for row in dictionary["ship-address-1"]:
        newRow = dictionary["ship-address-1"][row]
        rowIndex = row
        infoAddress2 = dictionary["ship-address-2"][row]
        if str(infoAddress2) == "nan":
            infoAddress2 = ""
        dictionary["ship-address-1"][row] = newRow + " " + str(infoAddress2)
    dictionary.pop("ship-address-2")
    addNewColumns(dictionary)


def main():
    rawStringPath = input("Introduzca la direccion del archivo: ")
    rawStringPath = rawStringPath.replace('"', "")
    convertToCSV(r"{s}".format(s=rawStringPath))
    initial_dict = readCSV(
        r"{s}/output.csv".format(s=os.getcwd()), r"{s}".format(s=rawStringPath)
    )
    formatCSV(initial_dict)
    filename = "{d}.xlsx".format(d=datetime.datetime.now().strftime("%d-%m-%Y"))
    os.system("cls")
    print(
        Fore.GREEN
        + 'Proceso finalizado.\nArchivo generado: "{f}"'.format(f=filename)
        + Fore.RESET
    )


main()
