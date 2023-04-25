package org.example
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.FileInputStream
import java.io.FileOutputStream
import java.time.LocalTime


/*
fun main(args: Array<String>) {
    val xlWb = XSSFWorkbook()
    val xlWs = xlWb.createSheet()
        xlWs.createRow(0).createCell(0).setCellValue("Привет жабам!")
    val output = FileOutputStream("./test.xlsx")
    xlWb.write(output)
    xlWb.close()
}
 */
fun main(args: Array<String>) {
    val input = FileInputStream("./Протокол_разбора_видео_озерная_лягушка_2022_метаданные_01_04_1.xlsx")
    val xlWb = WorkbookFactory.create(input)
    val xlWs = xlWb.getSheetAt(6)
    var end = false
    var rowForEvents = 0
    var rowForResults = 0
    var count = 0
    while (!end) {

        val round = xlWs.getRow(0+rowForEvents).getCell(7)
        val fishNumber = xlWs.getRow(0+rowForEvents).getCell(7)
        //println(round)
        //println(xlWs.getRow(8).getCell(1).toString())
        var n = 1
        var time = mutableListOf<LocalTime>()
        // Формирование списка времени события
        time.add(xlWs.getRow(7+rowForEvents).getCell(2).localDateTimeCellValue.toLocalTime())
        while (xlWs.getRow(7 + n+rowForEvents).getCell(1).toString() == "") {

            var timeCurr = xlWs.getRow(7 + n+rowForEvents).getCell(2).localDateTimeCellValue
            time.add(timeCurr.toLocalTime())
            //println(time)
            n++
        }
        // Формирование списка времени события в секундах
        var timeInSec = mutableListOf<Int>()
        for (date in time) {
            timeInSec.add(date.hour.toInt() * 60 + date.minute.toInt())
        }
        // Формирование списка времени события в секундах с нулевой секунды
        var timeSinceStart = mutableListOf<Int>()
        for (time in timeInSec) {
            timeSinceStart.add(time - timeInSec[0])
        }

        // Формирование списка кода события
        n = 1
        var eventCode = mutableListOf<String>()
        for (time in timeInSec) {
            eventCode.add(xlWs.getRow(6 + n+rowForEvents).getCell(5).toString())
            n++
        }

        // Формирование списка укусов
        n = 1
        var biteCode = mutableListOf<String>()
        for (time in timeInSec) {
            biteCode.add(xlWs.getRow(6 + n+rowForEvents).getCell(6).toString())
            n++
        }

        // фиксируем раунд
        //xlWs.getRow(7+rowForResults).getCell(7).setCellValue(xlWs.getRow(0+rowForEvents).getCell(6).toString())

        // фиксируем рыбу
        //xlWs.getRow(7+rowForResults).getCell(8).setCellValue(xlWs.getRow(1+rowForEvents).getCell(6).toString())


        // подcчет числа укусов в первые 5 секунд
        n = 0
        var bite = 0
        while (timeSinceStart[n] <= 5) n++

        for (i in 0 until n)
            if (eventCode[i] == "1.0") bite++
        xlWs.getRow(7+rowForResults).getCell(10).setCellValue(bite.toString())

        //Подсчет укусов в оставшееся время.
        var biteleft = 0
        for (i in n until eventCode.size)
            if (eventCode[i] == "1.0") biteleft++
        xlWs.getRow(7+rowForResults).getCell(11).setCellValue(biteleft.toString())

        //Подсчет укусов всего
        xlWs.getRow(7+rowForResults).getCell(12).setCellValue((biteleft + bite).toString())

        //Место первой атаки
        n = 0
        while (xlWs.getRow(8+rowForEvents+n).getCell(6).toString() == "") n++
        //var firstAttack: String = xlWs.getRow(8+rowForEvents).getCell(6)?.toString()?:""
        //xlWs.getRow(7+rowForResults).getCell(13).setCellValue(if (firstAttack != "")firstAttack.toFloat().toInt().toString() else firstAttack)
        xlWs.getRow(7+rowForResults).getCell(13).setCellValue(xlWs.getRow(8+rowForEvents+n).getCell(6).toString().toFloat().toInt().toString())
        // Место укуса 5 сек
        n = 0
        var bitePlaceFiveSec: String = ""
        while (timeSinceStart[n] <= 5) n++
        for (i in 0 until n)
            if (eventCode[i] == "1.0") bitePlaceFiveSec += "${biteCode[i].toFloat().toInt()}, "
        xlWs.getRow(7+rowForResults).getCell(14).setCellValue(bitePlaceFiveSec)

        // Место укуса оставшееся
        var bitePlaceAfterFiveSec: String = ""
        for (i in n until eventCode.size)
            if (eventCode[i] == "1.0") bitePlaceAfterFiveSec += "${biteCode[i].toFloat().toInt()}, "
        xlWs.getRow(7+rowForResults).getCell(15).setCellValue(bitePlaceAfterFiveSec)

        // Общеее место укуса
        xlWs.getRow(7+rowForResults).getCell(16).setCellValue(bitePlaceFiveSec + bitePlaceAfterFiveSec)

        // Число удержаний в первые 5 секунд
        n = 0
        var holdingFiveSec = 0
        while (timeSinceStart[n] <= 5) n++
        for (i in 0 until n)
            if (eventCode[i] == "2.0") holdingFiveSec++
        xlWs.getRow(7+rowForResults).getCell(17).setCellValue(holdingFiveSec.toString())

        // Число оставшихся удержаний
        var holdingAfterFiveSec = 0
        for (i in n until eventCode.size)
            if (eventCode[i] == "2.0") holdingAfterFiveSec++
        xlWs.getRow(7+rowForResults).getCell(18).setCellValue(holdingAfterFiveSec.toString())

        // Число общее удержаний
        xlWs.getRow(7+rowForResults).getCell(19).setCellValue((holdingFiveSec + holdingAfterFiveSec).toString())

        // Время удержаний по эпизодам и общее
        var holdStart = 0
        var holdSecSumm = 0
        var holdSecString = ""
        for (i in 0 until eventCode.size) {
            if (eventCode[i] == "2.0") holdStart = timeInSec[i]
            else if (eventCode[i] == "3.0") {
                holdSecSumm += timeInSec[i] - holdStart
                holdSecString += "${timeInSec[i] - holdStart},"
            }
        }
        xlWs.getRow(7+rowForResults).getCell(20).setCellValue(holdSecString)
        xlWs.getRow(7+rowForResults).getCell(21).setCellValue(holdSecSumm.toString())

        // Место удержания 5 секунд
        n = 0
        var holdLocationFiveSec: String = ""
        while (timeSinceStart[n] <= 5) n++
        for (i in 0 until n)
            if (eventCode[i] == "2.0") holdLocationFiveSec += "${biteCode[i].toFloat().toInt()}, "
        xlWs.getRow(7+rowForResults).getCell(22).setCellValue(holdLocationFiveSec)

        // Место удержания после 5 секунд
        var holdLocationAfterFiveSec: String = ""
        for (i in n until eventCode.size)
            if (eventCode[i] == "2.0") holdLocationAfterFiveSec += "${biteCode[i].toFloat().toInt()}, "
        xlWs.getRow(7+rowForResults).getCell(23).setCellValue(holdLocationAfterFiveSec)

        // Общее удержание
        xlWs.getRow(7+rowForResults).getCell(24).setCellValue(holdLocationFiveSec + holdLocationAfterFiveSec)

        //Количество жеваний 5 сек
        n = 0
        var chewingFiveSec = 0
        while (timeSinceStart[n] <= 5) n++
        for (i in 0 until n)
            if (eventCode[i] == "4.0") chewingFiveSec++
        xlWs.getRow(7+rowForResults).getCell(25).setCellValue(chewingFiveSec.toString())

        //Количество жеваний после 5 сек
        var chewingAfterFiveSec = 0
        for (i in n until eventCode.size)
            if (eventCode[i] == "4.0") chewingAfterFiveSec++
        xlWs.getRow(7+rowForResults).getCell(26).setCellValue(chewingAfterFiveSec.toString())

        //Количество жеваний общее
        xlWs.getRow(7+rowForResults).getCell(27).setCellValue((chewingFiveSec + chewingAfterFiveSec).toString())

        //Количество выплевываний 5 сек
        n = 0
        var spittingFiveSec = 0
        while (timeSinceStart[n] <= 5) n++
        for (i in 0 until n)
            if (eventCode[i] == "5.0") spittingFiveSec++
        xlWs.getRow(7+rowForResults).getCell(28).setCellValue(spittingFiveSec.toString())

        //Количество выплевываний после 5 сек
        var spittingAfterFiveSec = 0
        for (i in n until eventCode.size)
            if (eventCode[i] == "5.0") spittingAfterFiveSec++
        xlWs.getRow(7+rowForResults).getCell(29).setCellValue(spittingAfterFiveSec.toString())

        // общее выплевывание
        xlWs.getRow(7+rowForResults).getCell(30).setCellValue((spittingAfterFiveSec + spittingFiveSec).toString())

        // Время жевания
        var chewStart = 0
        var chewSecSumm = 0
        var chewSecString = ""
        for (i in 0 until eventCode.size) {
            if (eventCode[i] == "4.0") chewStart = timeInSec[i]
            else if (eventCode[i] == "5.0") {
                var timeChewing = timeInSec[i] - chewStart
                if (timeChewing == 0) timeChewing = 1
                chewSecSumm += timeChewing
                chewSecString += "$timeChewing,"
            }
        }
//        if (chewSecSumm == 0)  chewSecSumm = 1
//        if (chewSecString == "") chewSecString = "1"
        xlWs.getRow(7+rowForResults).getCell(31).setCellValue(chewSecString)
        xlWs.getRow(7+rowForResults).getCell(32).setCellValue(if (chewSecSumm == 0) "" else chewSecSumm.toString())

        // Заглатывание 5 сек
        n = 0
        var swallowFiveSec = 0
        while (timeSinceStart[n] <= 5) n++
        for (i in 0 until n)
            if (eventCode[i] == "6.0") swallowFiveSec++
        xlWs.getRow(7+rowForResults).getCell(33).setCellValue(swallowFiveSec.toString())

        //Количество Заглатываний после 5 сек
        var swallowAfterFiveSec = 0
        for (i in n until eventCode.size)
            if (eventCode[i] == "6.0") swallowAfterFiveSec++
        xlWs.getRow(7+rowForResults).getCell(34).setCellValue(swallowAfterFiveSec.toString())

        // общее Заглатывание
        xlWs.getRow(7+rowForResults).getCell(35).setCellValue((swallowFiveSec + swallowAfterFiveSec).toString())

        // Количество отрыгиваний 5 сек
        n = 0
        var burpFiveSec = 0
        while (timeSinceStart[n] <= 5) n++
        for (i in 0 until n)
            if (eventCode[i] == "7.0") burpFiveSec++
        xlWs.getRow(7+rowForResults).getCell(36).setCellValue(burpFiveSec.toString())

        // Отрыгивание оставшееся

        var burpAfterFiveSec = 0
        for (i in n until eventCode.size)
            if (eventCode[i] == "7.0") burpAfterFiveSec++
        xlWs.getRow(7+rowForResults).getCell(37).setCellValue(burpAfterFiveSec.toString())

        // общее отрыгивание
        xlWs.getRow(7+rowForResults).getCell(38).setCellValue((burpFiveSec + burpAfterFiveSec).toString())

        // время заглатывания
        var swallowStart = 0
        var swallowSecSumm = 0
        var swallowSecString = ""
        var swallowTrue = false
        for (i in 0 until eventCode.size) {
            if (eventCode[i] == "6.0") {
                swallowStart = timeInSec[i]
                swallowTrue = true
            }
            else if (eventCode[i] == "7.0") {
                var swallowTime = timeInSec[i] - swallowStart
                if (swallowTime == 0) swallowTime = 1
                swallowSecSumm += swallowTime
                swallowSecString += "$swallowTime,"
                swallowTrue = false
            }
            else if (eventCode[i] == "8.0" && swallowTrue) {
                var swallowTime = timeInSec[i] - swallowStart
                if (swallowTime == 0) swallowTime = 1
                swallowSecSumm += swallowTime
                swallowSecString += "$swallowTime,"
            }
        }
//        if (swallowSecSumm == 0)  swallowSecSumm = 1
//        if (swallowSecString == "") swallowSecString = "1"
        xlWs.getRow(7+rowForResults).getCell(39).setCellValue(swallowSecString)
        xlWs.getRow(7+rowForResults).getCell(40).setCellValue(if (swallowSecSumm==0) "" else swallowSecSumm.toString())
//        Число атак 5  сек
        n = 0
        var attackFiveSec = 0
        while (timeSinceStart[n] <= 5) n++
        for (i in 0 until n)
            when (eventCode[i]) {
                "1.0", "2.0","4.0","5.0","6.0","7.0" -> attackFiveSec++
            }
        xlWs.getRow(7+rowForResults).getCell(41).setCellValue(attackFiveSec.toString())

        //  Число атак после 5 сек.
        var attackAfterFiveSec = 0
        for (i in n until eventCode.size)
            when (eventCode[i]) {
                "1.0", "2.0","4.0","5.0","6.0","7.0" -> attackAfterFiveSec ++
            }
        xlWs.getRow(7+rowForResults).getCell(42).setCellValue(attackAfterFiveSec.toString())

        //  Число атак общее
        xlWs.getRow(7+rowForResults).getCell(43).setCellValue((attackAfterFiveSec+attackFiveSec).toString())

        // Место атаки 5 сек
        n = 0
        var attackPlaceFiveSec: String = ""
        while (timeSinceStart[n] <= 5) n++
        for (i in 0 until n)
            when (eventCode[i]) {
                "1.0", "2.0" -> attackPlaceFiveSec += "${biteCode[i].toFloat().toInt()}, "
            }
        xlWs.getRow(7+rowForResults).getCell(44).setCellValue(attackPlaceFiveSec)

        // Место атаки после 5 сек.
        var attackPlaceAfterFiveSec: String = ""
        for (i in n until eventCode.size)
            when (eventCode[i]) {
                "1.0", "2.0" -> attackPlaceAfterFiveSec += "${biteCode[i].toFloat().toInt()}, "
            }
        xlWs.getRow(7+rowForResults).getCell(45).setCellValue(attackPlaceAfterFiveSec)

        //Атака вся
        xlWs.getRow(7+rowForResults).getCell(46).setCellValue(attackPlaceFiveSec + attackPlaceAfterFiveSec)
        val output = FileOutputStream("./Протокол_разбора_видео_озерная_лягушка_2022_метаданные_01_04_12.xlsx")
        xlWb.write(output)


        rowForEvents += time.size
        rowForResults++
        if (xlWs.getRow(rowForEvents+1).getCell(2) == null) end = true

    println("Успешно обработана строка $count")
        count++
    }

    xlWb.close()
}