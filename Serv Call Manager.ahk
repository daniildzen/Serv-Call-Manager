#Requires AutoHotkey v2.0
; ========================
; === ГЛОБАЛЬНЫЕ НАСТРОЙКИ ===
; ========================
#SingleInstance Force
SendMode "Input"
SetWorkingDir A_ScriptDir

; =================
; === РЕСУРСЫ ===
; =================
resources(name) {
    return A_ScriptDir "\resources\" name
}

; =====================
; === ИНИЦИАЛИЗАЦИЯ ===
; =====================
Initialize() {
    ; Загрузочный экран
    splashGui := CreateSplashGUI()
    Sleep 1500
    splashGui.Destroy()
    
    ; GUI элементов управления
    global pauseGui := CreatePauseGUI()
    global infoGui := CreateInfoGUI()
    global abbrevGui := CreateAbbrevGUI()
    
    ; Позиционирование GUI
    global offsetBottom := 40 ; Отступ от нижнего края

    infoGui.Show("Hide")
    infoGui.GetPos(,,, &iH)
    global infoH := iH

    abbrevGui.Show("Hide")
    abbrevGui.GetPos(,,, &aH)
    global abbrevH := aH

    pauseGui.Show("Hide")
    pauseGui.GetPos(,,, &pH)
    global pauseH := pH
    
    infoGui.Show("x0 y" (A_ScreenHeight - infoH - offsetBottom))
}

; =====================
; === ГРАФИЧЕСКИЕ ИНТЕРФЕЙСЫ ===
; =====================
CreateSplashGUI() {
    splash := Gui("+AlwaysOnTop -Caption -Border +ToolWindow")
    splash.BackColor := "0x222222"
    splash.MarginX := 0, splash.MarginY := 0
    
    imagePath := resources("logo.png")
    if FileExist(imagePath) {
        pic := splash.Add("Picture", "w350 h350", imagePath)
        pic.OnEvent("Click", (*) => ExitApp())
    } else {
        splash.Add("Text", "w350 h350 Center 0x200 cWhite", "Serv Call Manager")
    }
    
    progress := splash.Add("Progress", "w350 h8 Range0-1500", 0)
    startTime := A_TickCount
    
    SetTimer(updateProgress, 100)
    updateProgress() {
        elapsed := A_TickCount - startTime
        if (elapsed >= 500) {
            progress.Value := 500
            SetTimer(, 0)
            return
        }
        progress.Value := elapsed
    }
    
    splash.Show("Center")
    splash.Title := "Загрузка"
    return splash
}

CreatePauseGUI() {
    guiObj := Gui("+AlwaysOnTop -Caption -Border +ToolWindow")
    guiObj.BackColor := "0x2C2C2C"
    guiObj.SetFont("s12 cWhite Bold", "Segoe UI")
    guiObj.Add("Text", "w200 h50 Center 0x200", "⏸ Скрипт на паузе")
    return guiObj
}

CreateInfoGUI() {
    guiObj := Gui("+AlwaysOnTop -Caption -Border +ToolWindow")
    guiObj.BackColor := "0x1E1E1E"
    guiObj.SetFont("s10 cCCCCCC", "Consolas")
    guiObj.MarginX := 10, guiObj.MarginY := 10
    
    text := "
    (LTrim
        [F6]  Генератор номера ВТБ
        [F8]  Генератор номера АБ
        [F7]  Шаблон комментария
        [F9]  Перезапуск скрипта
        [F10] Пауза/возобновить
        [F11] Скрыть/показать подсказки
        
        [Ctrl+F6] Письмо приостановки
        [Ctrl+F7] Комментарий (одиночный)
        [Ctrl+F9] Список сокращений
    )"
    
    guiObj.Add("Text", "w280 h150", text)
    return guiObj
}

CreateAbbrevGUI() {
    guiObj := Gui("+AlwaysOnTop -Caption -Border +ToolWindow")
    guiObj.BackColor := "0x1E1E1E"
    guiObj.SetFont("s9 cCCCCCC", "Consolas")
    guiObj.MarginX := 10, guiObj.MarginY := 10
    
    text := "
    (LTrim
        [верн, согл]            Адрес верный + график
        [нд]                    Не дозвон
        [ао]                    Автоответчик
        [сброс, сбр]            Сброс вызова
        [тишина, тихо, молчат]  Тишина в трубке
        [нетсп, нк]             Не принадлежит ТСП
        [не ак, не акт]         Заявка не актуальна
        [раб вост]              Терминал восстановлен
        [по доп]                По доп. номеру
        [по осн]                По основному номеру
        [неверн]                Адрес неверный
        [247, кругл]            Круглосуточно
        [не сущ]                Номер не существует
        
        [Ctrl+F9] ◀ Назад
    )"
    
    guiObj.Add("Text", "w320 h210", text)
    return guiObj
}

; =====================
; === ОСНОВНОЙ КЛАСС ===
; =====================
class CallManager {
    static TIDCopy(bank) {
        A_Clipboard := ""
        Loop {
            ToolTip("[" bank "]: Скопируйте ID терминала")
            SetTimer(() => ToolTip(), -2000)
            if A_Clipboard
                return SubStr(A_Clipboard, -6)
        }
    }

    static REQCopy(bank) {
        A_Clipboard := ""
        Loop {
            ToolTip("[" bank "]: Скопируйте ID заявки")
            SetTimer(() => ToolTip(), -2000)
            if A_Clipboard
                return SubStr(A_Clipboard, -4)
        }
    }

    static ActivateInfinity() {
        if WinExist("ahk_exe Cx.Client.exe") {
            WinActivate
            Send "^v{Enter}"
        } else {
            MsgBox("Infinity не запущен!`nЗапустите программу для работы функции", "Ошибка", "Icon!")
        }
    }

    static NumberGenerator() {
        bank := "ВТБ"
        tid := this.TIDCopy(bank)
        req := this.REQCopy(bank)
        A_Clipboard := "88007007321,9," tid "," req ",1"
        this.ActivateInfinity()
        ToolTip("[" bank "]: Номер сгенерирован", , 2)
        SetTimer(() => ToolTip(), -2000)
    }

    static ABNumberGenerator() {
        bank := "АБ"
        tid := this.TIDCopy(bank)
        req := this.REQCopy(bank)
        A_Clipboard := "88007004891,9," tid "," req ",1"
        this.ActivateInfinity()
        ToolTip("[" bank "]: Номер сгенерирован", , 2)
        SetTimer(() => ToolTip(), -2000)
    }

    static CommentGenerator() {
        now := A_Now
        time1 := DateAdd(now, 30, "Minutes")
        time2 := DateAdd(now, 60, "Minutes")
        
        SendInput FormatTime(now, "dd.MM`n    HH:mm`n    ")
        SendInput FormatTime(time1, "HH:mm`n    ")
        SendInput FormatTime(time2, "HH:mm")
    }

    static SingleCommentGenerator() {
        SendInput FormatTime(, "dd.MM: HH:mm - ")
    }
}

; =====================
; === ГОРЯЧИЕ КЛАВИШИ ===
; =====================
#SuspendExempt
Initialize()

; Основные функции
F6:: CallManager.NumberGenerator()
F8:: CallManager.ABNumberGenerator()
F7:: CallManager.CommentGenerator()
^F7:: CallManager.SingleCommentGenerator()
^F6:: Run(resources("letter.msg"))

; Управление скриптом
F9:: Reload
F10:: TogglePause()
F11:: ToggleInfoDisplay()
^F9:: ToggleAbbrevDisplay()

; Текстовые сокращения (полностью восстановленные)
::верн::  { 
SendText("Адрес верный. График: с `nОриентир: ")
}
::dthy::  { 
SendText("Адрес верный. График: с `nОриентир: ")
}
::согл::  { 
SendText("Адрес верный. График: с `nОриентир: ")
}
::cjuk::  { 
SendText("Адрес верный. График: с `nОриентир: ")
}

::неверн::{ 
SendText("Адрес не верный. Фактический адрес . График: с `nОриентир: ")
}
::ytdthy::{ 
SendText("Адрес не верный. Фактический адрес . График: с `nОриентир: ")
}

::247::   { 
SendText("круглосуточно")
}
::кругл:: { 
SendText("круглосуточно")
}
::rheuk:: { 
SendText("круглосуточно")
}

::нд::    { 
SendText("Не удалось связаться с клиентом.")
}
::yl::    { 
SendText("Не удалось связаться с клиентом.")
}

::ао::    {
SendText("Не удалось связаться с клиентом: Автоответчик.")
}
::fj::    {
SendText("Не удалось связаться с клиентом: Автоответчик.")
}

::сброс:: {
SendText("Не удалось связаться с клиентом: Клиент сбросил звонок.")
}
::c,hjc:: {
SendText("Не удалось связаться с клиентом: Клиент сбросил звонок.")
}
::сбр::   {
SendText("Не удалось связаться с клиентом: Клиент сбросил звонок.")
}
::c,h::   {
SendText("Не удалось связаться с клиентом: Клиент сбросил звонок.")
}

::тишина::{
SendText("Не удалось связаться с клиентом: Тишина в трубке.")
}
::тихо::  {
SendText("Не удалось связаться с клиентом: Тишина в трубке.")
}
::молчат::{
SendText("Не удалось связаться с клиентом: Тишина в трубке.")
}
::nbibyf::{
SendText("Не удалось связаться с клиентом: Тишина в трубке.")
}
::nb[j::  {
SendText("Не удалось связаться с клиентом: Тишина в трубке.")
}
::vjkxfn::{
SendText("Не удалось связаться с клиентом: Тишина в трубке.")
}

::нетсп:: {
SendText("Контактное лицо не принадлежит ТСП. Просьба предоставить доп. контакт.")
}
::нк::    {
SendText("Контактное лицо не принадлежит ТСП. Просьба предоставить доп. контакт.")
}
::ytncg:: {
SendText("Контактное лицо не принадлежит ТСП. Просьба предоставить доп. контакт.")
}
::yr::    {
SendText("Контактное лицо не принадлежит ТСП. Просьба предоставить доп. контакт.")
}

::по доп::{
SendText("По доп. номеру: ")
}
::gj ljg::{
SendText("По доп. номеру: ")
}

::по осн::{
SendText("По основому номеру: ")
}
::gj jcy::{
SendText("По основому номеру: ")
}

::не ак:: {
SendText("Со слов клиента заявка не актуальна.")
}
::не акт::{
SendText("Со слов клиента заявка не актуальна.")
}
::yt fr:: {
SendText("Со слов клиента заявка не актуальна.")
}
::yt frn::{
SendText("Со слов клиента заявка не актуальна.")
}

::раб вост::{
SendText("Работоспособность терминала восстановлена.")
}
::hf, djcn::{
SendText("Работоспособность терминала восстановлена.")
}

::не сущ:: {
SendText("Номер не существует. Необходим доп. номер.")
}
::yt ceo:: {
SendText("Номер не существует. Необходим доп. номер.")
}
#SuspendExempt False

; =====================
; === ФУНКЦИИ УПРАВЛЕНИЯ ===
; =====================
TogglePause() {
    static paused := false
    paused := !paused
    
    if paused {
        Suspend 1
        infoGui.Hide()
        abbrevGui.Hide()
        pauseGui.Show("x0 y" (A_ScreenHeight - pauseH - offsetBottom))
    } else {
        Suspend 0
        pauseGui.Hide()
        infoGui.Show("x0 y" (A_ScreenHeight - infoH - offsetBottom))
    }
}

ToggleInfoDisplay() {
    static visible := true
    visible := !visible
    if visible {
        infoGui.Show("x0 y" (A_ScreenHeight - infoH - offsetBottom))
    } else {
        infoGui.Hide()
        abbrevGui.Hide()
    }
}

ToggleAbbrevDisplay() {
    static abbrevVisible := false
    abbrevVisible := !abbrevVisible
    
    if abbrevVisible {
        infoGui.Hide()
        abbrevGui.Show("x0 y" (A_ScreenHeight - abbrevH - offsetBottom))
    } else {
        abbrevGui.Hide()
        infoGui.Show("x0 y" (A_ScreenHeight - infoH - offsetBottom))
    }
}