#Requires AutoHotkey v2.0

; Загрузочный GUI
myGui := Gui("+AlwaysOnTop -Caption -Border")  ; Окно будет поверх всех других
imagepath := A_ScriptDir . "\resources\logo.png"

if FileExist(imagePath) {
    myGui.Add("Picture",, imagePath)
} else {
    myGui.Add("Text", "w300 h150 Center 0x200", "Serv Call Manager loading...")
}

myGui.Title := "Image Overlay"
myGui.Show()
sleep(1500)
myGui.Destroy()

; GUI для паузы
pauseGui := Gui("+AlwaysOnTop -Caption -Border")
pauseGui.BackColor := "EEEEEE"
pauseGui.SetFont("s10", "Arial")
pauseGui.Add("Text", "w300", "Скрипт на паузе")

; GUI для хоткеев
infoGui := Gui("+AlwaysOnTop -Caption -Border")
infoGui.BackColor := "EEEEEE"
infoGui.SetFont("s10", "Arial")
infoGui.Add("Text", "w300", "
(Join`n
    F6 - Сгенерировать номер ВТБ
    F8 - Сгенерировать номер АБ
    ----
    F7 - Сгенерировать комментарий 
    Ctrl + F7 - Комментарий - Одиночный звонок
    Ctrl + F6 - Письмо для приостановки
    ---
    F9 - Перезагрузить программу
    Ctrl + F9 - Список сокращений
    F10 - Пауза программы
    F11 - Скрыть подсказку
)")

; GUI для аббревиатур
abbrevGui := Gui("+AlwaysOnTop -Caption -Border")
abbrevGui.BackColor := "EEEEEE"
abbrevGui.SetFont("s10", "Arial")
abbrevGui.Add("Text", "w350", "
(
    [верн, согл]: Адрес верный. График: с <...> Ориентир:
    [24-7, 247, кругл]: круглосуточно
    [нд]: Не удалось связаться с клиентом.
    [ао]: Автоответчик.
    [сброс, сбр]: Клиент сбросил звонок.
    [тишина, тихо, молчат]: Тишина в трубке.
    [нетсп, нк]: Контактное лицо не принадлежит ТСП.
    [не ак, не акт]: Со слов клиента заявка не актуальна.
    [раб вост]: Работоспособность терминала восстановлена. 
    [по доп]: По доп. номеру:
    [по осн]: По основому номеру:
    [не сущ]: Номер не существует. Необходим доп. номер.
    [неверн]: Адрес не верный.

    Ctrl+F9 - назад
)")

global infoGuiPos := 230
global abbrevGuiPos := 303

infoGui.Show("x0 y" A_ScreenHeight-infoGuiPos)

class CallManager{

    TIDCopy(bank){
        A_Clipboard := ""
        Loop{
            ToolTip("[" bank "]: Скопируйте ID терминала")
            SetTimer(() => ToolTip(), -2000)

            if(A_Clipboard){
                return terminal := SubStr(A_Clipboard, -6)
            }
        }
    }

    REQCopy(bank){
        A_Clipboard := ""
        Loop{
            ToolTip("[" bank "]: Скопируйте ID заявки")
            SetTimer(() => ToolTip(), -2000)
            if(A_Clipboard){
                return req := SubStr(A_Clipboard, -4)
            }
        }
    }

    ActivateInfitity(){
        if(WinExist("ahk_exe Cx.Client.exe")){
            WinActivate("ahk_exe Cx.Client.exe")
            Send("^v")
            Send("{Enter}")
        }
        else{
            MsgBox("Infinity не запущен! Пожалуйста, запустите Infinity чтобы работала функция автоматических звонков.", "Не обнаружен Infinity Call-center")
            return
        }
    }

    NumberGenerator(){
        bank := "ВТБ"
        terminal := this.TIDCopy(bank)
        req := this.REQCopy(bank)

        A_Clipboard := "88007007321,9," terminal "," req ",1"
        ToolTip("[" bank "]: Номер успешно сгенерирован (" A_Clipboard ")")
        this.ActivateInfitity()
    }

    ABNumberGenerator(){
        bank := "АБ"
        terminal := this.TIDCopy(bank)
        req := this.REQCopy(bank)

        A_Clipboard := "88007004891,9," terminal "," req ",1"
        ToolTip("[" bank "]: Номер успешно сгенерирован (" A_Clipboard ")")
        this.ActivateInfitity()
    }

    CommentGenerator(){
        CurrentDate := FormatTime(, "dd.MM")
        CurrentTime := FormatTime(, "HH:mm")

        TimePlus1 := AddMinutes(CurrentTime, 30)
        TimePlus2 := AddMinutes(CurrentTime, 60)

        Output := CurrentDate "`n    " CurrentTime "`n    " TimePlus1 "`n    " TimePlus2
        SendInput(Output)

        ; Функция для добавления минут к времени в формате HH:mm
        AddMinutes(Time, MinutesToAdd) {
        ; Разбираем время на часы и минуты
        TimeParts := StrSplit(Time, ":")
        Hours := Integer(TimeParts[1])
        Minutes := Integer(TimeParts[2])

        ; Преобразуем в общее количество минут
        TotalMinutes := Hours * 60 + Minutes + MinutesToAdd

        ; Вычисляем новые часы и минуты (с учетом перехода через 24 часа)
        NewHours := Mod(Floor(TotalMinutes / 60), 24)
        NewMinutes := Mod(TotalMinutes, 60)

        ; Форматируем с ведущими нулями
        NewHours := Format("{:02}", NewHours)
        NewMinutes := Format("{:02}", NewMinutes)

        return NewHours ":" NewMinutes
        }
    }

    SingleCommentGenerator(){
        CurrentDate := FormatTime(, "dd.MM")
        CurrentTime := FormatTime(, "HH:mm")
        Output := CurrentDate ": " CurrentTime " - " 
        SendInput(Output)
    }
}

; Сокращения

F6::{
    manager := CallManager()
    manager.NumberGenerator()
}

^F6::{
    letterpath := A_ScriptDir . "\resources\letter.msg"
    Run(letterpath)
}

F7::{
    manager := CallManager()
    manager.CommentGenerator()
}

^F7::{
    manager := CallManager()
    manager.SingleCommentGenerator()
}

F8::{
    manager := CallManager()
    manager.ABNumberGenerator()
}


::верн:: {
    Send("Адрес верный. График: с `nОриентир: ")
}
::согл:: {
    Send("Адрес верный. График: с `nОриентир: ")
}
::dthy:: {
    Send("Адрес верный. График: с `nОриентир: ")
}
::cjuk:: {
    Send("Адрес верный. График: с `nОриентир: ")
}

::неверн:: {
    Send("Адрес не верный. Фактический адрес . График: с `nОриентир: ")
}
::ytdthy:: {
    Send("Адрес не верный. Фактический адрес . График: с `nОриентир: ")
}

::247::круглосуточно
::кругл::круглосуточно
::rheuk::круглосуточно

::нд::Не удалось связаться с клиентом.
::yl::Не удалось связаться с клиентом.

::ао::Не удалось связаться с клиентом: Автоответчик.
::fj::Не удалось связаться с клиентом: Автоответчик.

::сброс::Не удалось связаться с клиентом: Клиент сбросил звонок.
::c,hjc::Не удалось связаться с клиентом: Клиент сбросил звонок.
::сбр::Не удалось связаться с клиентом: Клиент сбросил звонок.
::c,h::Не удалось связаться с клиентом: Клиент сбросил звонок.

::тишина::Не удалось связаться с клиентом: Тишина в трубке.
::тихо::Не удалось связаться с клиентом: Тишина в трубке.
::молчат::Не удалось связаться с клиентом: Тишина в трубке.
::nbibyf::Не удалось связаться с клиентом: Тишина в трубке.
::nb[j::Не удалось связаться с клиентом: Тишина в трубке.
::vjkxfn::Не удалось связаться с клиентом: Тишина в трубке.

::нетсп::Контактное лицо не принадлежит ТСП. Просьба предоставить доп. контакт.
::нк::Контактное лицо не принадлежит ТСП. Просьба предоставить доп. контакт.
::ytncg::Контактное лицо не принадлежит ТСП. Просьба предоставить доп. контакт.
::yr::Контактное лицо не принадлежит ТСП. Просьба предоставить доп. контакт.

::по доп::По доп. номеру: 
::gj ljg::По доп. номеру: 

::по осн::По основому номеру: 
::gj jcy::По основому номеру: 

::не ак::Со слов клиента заявка не актуальна.
::не акт::Со слов клиента заявка не актуальна.
::yt fr::Со слов клиента заявка не актуальна.
::yt frn::Со слов клиента заявка не актуальна.

::раб вост::Работоспособность терминала восстановлена.
::hf, djcn::Работоспособность терминала восстановлена.

::не сущ::Номер не существует. Необходим доп. номер.
::yt ceo::Номер не существует. Необходим доп. номер.

; Обработчик для Ctrl+F9
#SuspendExempt true
^F9::{
    static showAbbreviations := false
    
    if (showAbbreviations) {
        abbrevGui.Hide()
        infoGui.Show("x0 y" A_ScreenHeight - infoGuiPos)
    } else {
        infoGui.Hide()
        abbrevGui.Show("x0 y" A_ScreenHeight - abbrevGuiPos)
    }
    
    showAbbreviations := !showAbbreviations
}

F9:: Reload

F10::
{
    if(A_IsSuspended){
        Suspend(0)
        infoGui.Show("x0 y" A_ScreenHeight - infoGuiPos)
        pauseGui.Hide()
    }
    else{
        Suspend(1)
        infoGui.Hide()
        abbrevGui.Hide()
        pauseGui.Show("x0 y" A_ScreenHeight - 75)
    }
}

F11::{
    static showHint := false
    
    if (showHint) {
        infoGui.Show("x0 y" A_ScreenHeight - infoGuiPos)
    } else {
        infoGui.Hide()
        abbrevGui.Hide()
    }
    
    showHint := !showHint
}
#SuspendExempt false