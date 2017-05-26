import de.btobastian.javacord.Javacord
import de.btobastian.javacord.entities.Channel
import de.btobastian.javacord.listener.message.MessageCreateListener
import org.apache.poi.hssf.usermodel.*
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFClientAnchor
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFSimpleShape
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.jsoup.Jsoup
import java.io.BufferedInputStream
import java.io.IOException
import java.net.URL
import java.text.SimpleDateFormat
import java.time.Instant
import java.util.concurrent.Executors
import java.util.concurrent.TimeUnit


class Tuple<X, Y>(var item1: X, var item2: Y)
class NoteData(var x1: Int, var y1: Int, var x2: Int, var y2: Int)
class AnnouncementData(var title: String, var date: String, var text: String = "", var url: String = ""){
    override fun toString(): String {
        return "$title - ($date)\n" +
                if(text.isNotEmpty()) "$text\n" else "" +
                if(url.isNotEmpty()) "$url\n" else ""
    }
}

val schoolURL = "http://www.handasaim.co.il/"
val command = '$'
var FIRST_LINE = 0
val channels = mutableListOf<Tuple<Channel, String>>()
val times = arrayOf(
        "07:45-08:30", // 0
        "08:30-09:15", // 1
        "09:15-10:00", // 2
        //"10:00-10:15",// null
        "10:15-11:00", // 3
        "11:00-11:45", // 4
        //"11:45-12:10",// null
        "12:10-12:55", // 5
        "12:55-13:40", // 6
        //"13:40-13:50",// null
        "13:50-14:35", // 7
        "14:35-15:20", // 8
        //"15:20-15:25",// null
        "15:25-16:10", // 9
        "16:10-16:55" // 10
)

fun main(args: Array<String>) {
    val time = when(args.size){
        0 -> {
            println("use: SchooledBot token [time in hours between update checks, 0 for no checks, 1 by default]")
            return
        }
        2 -> args[2].toLong()
        else -> 1
    }

    val api = Javacord.getApi(args[0], true)
    api.connectBlocking()
    api.game = "${command}help for commands"

    api.registerListener(MessageCreateListener { _, message ->
        var message = message
        if(message.content.first() == command) {
            val arguments = message.content.split(" ")
            var out = ""

            when(arguments[0].removeRange(0, 1)){
                "schedule" ->
                        try {
                            message = message.reply("working...").get()

                            if (arguments.size > 1)
                                out = schedule(arguments[1])
                            else
                                out = "schedule is used like so:\n" + command + "schedule *class*\nwhile *class* can be \"יא1\" ,\"ט2\" and so on"
                        }catch (ioe: IOException){
                            ioe.printStackTrace()
                            out = "error getting schedule"
                        }catch (e: Exception){
                            e.printStackTrace()
                        }
                "announcements",
                "ann" ->
                    try {
                        message = message.reply("working...").get()

                        for (data in announcementData())
                            out += data.toString() + '\n'
                    }catch (ioe: IOException){
                        ioe.printStackTrace()
                        out = "error parsing announcements"
                    }catch (e: Exception){
                        e.printStackTrace()
                    }

                "notes" ->
                        try {
                            message = message.reply("working...").get()

                            out = getNotes()
                        }catch (e: Exception){
                            e.printStackTrace()
                            out = "error getting notes"
                        }
                "setUpdateChannel",
                "set" ->
                    if(!channels.any{ it.item1 === message.channelReceiver }){
                        channels.add(Tuple(message.channelReceiver, ""))
                        out = "This channel is now set for updates!"
                    } else
                        out = "This channel is already scheduled for updates!"

                "removeUpdateChannel",
                "remove" ->{
                    val pre = channels.size
                    for (channel in channels)
                        if(channel.item1 === message.channelReceiver){
                            channels.remove(channel)
                            out = "The channel was removed successfully!"
                            break
                        }
                    if(pre == channels.size)
                        out = "This channel is not listed for updates"
                }

                "help",
                "commands" ->
                    out = "list of commands:\n" +
                            "**${command}schedule** *class* - while *class* can be \"יא1\" ,\"ט2\" and so on\n" +
                            "**${command}announcements/ann**\n" +
                            "**${command}notes**\n" +
                            "**${command}setUpdateChannel/set**\n" +
                            "**${command}removeUpdateChannel/remove**"
                else ->
                    out = "no such command, type \"${command}help\" for list of commands"
            }

            println("RECEIVED:\n" + message.content)
            println("\nSENT:\n$out\n")

            if (message.author.isYourself) {
                message.delete()
                message.reply(out)
            } else
                message.reply(out)
        }
    })

    if(time != 0L)
        Executors.newScheduledThreadPool(1).scheduleAtFixedRate({
            try {
                println("Now: " + Instant.now())
                val ann = relevantAnn()

                channels.forEach {
                    if (ann != null && ann.url != it.item2) {
                        it.item1.sendMessage("SCHOOL SITE UPDATED!\n" + ann.url)
                        it.item2 = ann.url
                    }
                }
            } catch (e: Exception) {
                e.printStackTrace()
            }
        }, 0L, time, TimeUnit.HOURS)
}

fun announcementData(): MutableList<AnnouncementData>{
    val doc = Jsoup.connect(schoolURL).get()

    val newsHeadlines = doc.select("marquee > table > tbody > tr > td")
    val data = newsHeadlines.html().split(String.format("(?=%1\$s)", "<sup>").toRegex())

    val announcement = mutableListOf<AnnouncementData>()

    for(str in data){
        var document = Jsoup.parse(str)

        val dataStr = document.select("sup").html()
        val date = dataStr.substring(1, dataStr.length - 1)

        val title = document.select("b").html()

        var url = document.select("a").attr("href").replace(" ", "")

        document = Jsoup.parse(document.toString()
                .replace(document.select("sup").toString(), "")
                .replace(document.select("b").toString(), "")
                .replace(document.select("a").toString(), ""))

        val text = document.select("body").html()
                .replace("<br>", "\n").replace("(?m)^[ \t]*\r?\n".toRegex(), "")

        fun urlValid(url: String) =
            try {
                URL(url)
                true
            } catch (e: java.net.MalformedURLException) {
                false
            }

        if(!urlValid(url))
            if (urlValid(schoolURL + url))
                url = schoolURL + url
            else
                url = "The web page \"$url\" does not exist"

        announcement.add(AnnouncementData(title, date, text, url))
    }

    return announcement
}

fun schedule(curClass: String): String{
    val itemData = relevantAnn() ?: return "אין מערכת"

    val isX = itemData.url.contains(".xlsx")
    val url = URL(itemData.url)
    val connection = url.openConnection()
    connection.connect()

    val excelFile = BufferedInputStream(url.openStream(), 8192)

    val workbook: Workbook
    if (isX)
        workbook = XSSFWorkbook(excelFile)
    else
        workbook = HSSFWorkbook(excelFile)

    val sheet = workbook.getSheetAt(0)
    val formulaEvaluator = workbook.creationHelper.createFormulaEvaluator()

    val day = getCellAsString(sheet.getRow(0), 0, formulaEvaluator)

    if (getCellAsString(sheet.getRow(0), 1, formulaEvaluator).isEmpty())
        FIRST_LINE = 1

    val rowsCount = sheet.physicalNumberOfRows
    var maxCols = 0

    (0 until rowsCount).forEach{
        if (sheet.getRow(it).physicalNumberOfCells > maxCols)
            maxCols = sheet.getRow(it).physicalNumberOfCells
    }

    var classNum = -1

    val classes = mutableListOf<String>()
    (1 until maxCols).forEach{
        val cell = getCellAsString(sheet.getRow(FIRST_LINE), it, formulaEvaluator)
        classes.add(cell.split(" ")[0])
        if (classes[classes.size - 1] == curClass)
            classNum = it - 1
    }

    if(classNum == -1){
        var err = "no such class. please choose a class from the following options:\n"
        (0 until classes.size - 1).forEach { err += classes[it] + ", " }
        err += classes[classes.size - 1]
        return err
    }

    val schedule = mutableListOf<String>()
    (0 until rowsCount).forEach {
        val row = sheet.getRow(it)
        val cells = Array(maxCols - 1){""}

        (0 until row.physicalNumberOfCells).forEach {
            if (it != 0 && it > FIRST_LINE)
                cells[it - 1] = getCellAsString(row, it, formulaEvaluator)
        }

        if (it > FIRST_LINE)
            if (it != rowsCount - 1)
                schedule.add(cells[classNum])
    }

    while (schedule.size > 0) {
        if (schedule[schedule.size - 1] != "")
            break
        schedule.removeAt(schedule.size - 1)
    }

    var out = ""

    var i = 0
    while (i < schedule.size) {
        if (i == 0 && schedule[i] == "")
            i++
        out += "\n\n" + times[i] + '\n' + schedule[i]
        i++
    }
    if (out.isNotEmpty())
        return "המערכת לכיתה $curClass ליום $day היא:$out"
    return "אין מערכת"
}

fun getNotes(): String {
    val itemData = relevantAnn()
    if(itemData == null)
        return "אין פתקים"

    val isX = itemData.url.contains(".xlsx")

    val url = URL(itemData.url)
    url.openConnection().connect()

    val excelFile = BufferedInputStream(url.openStream(), 8192)

    val workbook: Workbook
    if (isX)
        workbook = XSSFWorkbook(excelFile)
    else
        workbook = HSSFWorkbook(excelFile)

    val sheet = workbook.getSheetAt(0)
    val formulaEvaluator = workbook.creationHelper.createFormulaEvaluator()

    if (getCellAsString(sheet.getRow(0), 1, formulaEvaluator).isEmpty())
        FIRST_LINE = 1

    val rowsCount = sheet.physicalNumberOfRows
    var maxCols = 0

    (0 until rowsCount).forEach {
        if (sheet.getRow(it).physicalNumberOfCells > maxCols)
            maxCols = sheet.getRow(it).physicalNumberOfCells
    }

    val classes = mutableListOf<String>()
    (1 until maxCols).forEach {
        classes.add(getCellAsString(sheet.getRow(FIRST_LINE), it, formulaEvaluator))
    }


    if ((if (isX)
        (sheet as XSSFSheet).drawingPatriarch else
        (sheet as HSSFSheet).drawingPatriarch) != null) {
        val children = if (isX)
            (sheet as XSSFSheet).drawingPatriarch.shapes
        else
            (sheet as HSSFSheet).drawingPatriarch.children

        val it = children.iterator()

        while (it.hasNext()) {
            val anchor: ClientAnchor
            val str: String
            if (isX) {
                val shape = it.next() as XSSFSimpleShape
                anchor = shape.anchor as XSSFClientAnchor
                str = shape.text
            } else {
                val shape = it.next() as HSSFSimpleShape
                anchor = shape.anchor as HSSFClientAnchor
                str = shape.string.string
            }

            val data = NoteData(anchor.col1.toInt(), anchor.row1,
                    anchor.col2.toInt(), anchor.row2)

            return "```" + str + "\n\n" + getNoteInfo(data, classes) + "```\n"
        }
    }
    return "אין פתקים"
}

fun getNoteInfo(data: NoteData, classes: MutableList<String>): String {
    var classSelect = "ההודעה מופיעה מתחת לכיתות: "

    var i = data.x1
    while (i < data.x2 && i < classes.size) {
        classSelect += classes[i - 1] + ", "
        i++
    }
    classSelect += classes[if (i < classes.size) data.x2 - 1 else classes.size - 1]


    return classSelect +
            if (data.y1 != data.y2)
                "\nובין השעות ${data.y1 - FIRST_LINE - 1} ל ${data.y2 - FIRST_LINE - 1}"
            else
                "\nבשעה " + (data.y1 - FIRST_LINE - 1)

}

fun relevantAnn(): AnnouncementData? {
    announcementData().forEach {
        if (it.url.contains("s3-eu-west-1.amazonaws.com/schooly/handasaim/news") &&
                it.title.contains("מערכת שעות") && (it.url.contains(".xls") || it.url.contains(".xlsx")))
            return it
    }

    return null
}

fun getCellAsString(row: Row, c: Int, formulaEvaluator: FormulaEvaluator): String {
    val cell = row.getCell(c)
    val cellValue = formulaEvaluator.evaluate(cell)
    return if (cellValue != null)
        when (cellValue.cellType) {
            Cell.CELL_TYPE_BOOLEAN -> cellValue.booleanValue.toString()
            Cell.CELL_TYPE_NUMERIC ->
                if (HSSFDateUtil.isCellDateFormatted(cell))
                    SimpleDateFormat("dd/MM/yy").format(HSSFDateUtil.getJavaDate(cellValue.numberValue))
                else
                    java.lang.Double.toString(cellValue.numberValue)
            Cell.CELL_TYPE_STRING -> cellValue.stringValue
            else -> cellValue.toString()
        }
    else ""
}