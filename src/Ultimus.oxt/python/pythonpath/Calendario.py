"""
Module for generating dynamic dialogs
i.e. dialogs that auto-adjust their layout based on contents and
external constraints

Copyright 2020 by Massimo Del Fedele
"""
# import os
# import inspect
from datetime import date
# import uno
# import unohelper

from com.sun.star.style.VerticalAlignment import MIDDLE as VA_MIDDLE

from com.sun.star.awt import Size
from com.sun.star.awt import XActionListener, XTextListener
from com.sun.star.task import XJobExecutor

from com.sun.star.awt import XTopWindowListener

from com.sun.star.util import MeasureUnit

# from LeenoConfig import Config
import LeenoUtils
import pyleeno as PL
# import LeenoDialogs as DLG
import Dialogs


MINBTNWIDTH = 100

def calendario():
    '''
    Mostra un calendario da cui selezionare la data e la restituisce
    in formato gg/mm/aaaa.
    '''
    # oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet
    x = PL.LeggiPosizioneCorrente()[0]
    y = PL.LeggiPosizioneCorrente()[1]
    testo = pickDate()
    lst = str(testo).split('-')
    try:
        testo = lst[2] + '/' + lst[1] + '/' + lst[0]
    except:
        pass

    return testo

def pickDate(curDate=None):
    '''
    Allow to pick a date from a calendar
    '''
    if curDate is None:
        curDate = date.today()

    def rgb(r, g, b):
        return 256*256*r + 256*g + b


    btnWidth, btnHeight = Dialogs.getButtonSize('<<')
    dateWidth, dummy = Dialogs.getTextBox('88 SETTEMBRE 8888XX')

    workdaysBkColor = rgb(224, 224, 224)
    workdaysFgColor = rgb(0, 0, 0)
    holydaysBkColor = rgb(255, 153, 153)
    holydaysFgColor = rgb(0, 0, 0)

    # create daynames list with spacers
    dayNamesLabels = [Dialogs.FixedText(
        Text=LeenoUtils.DAYNAMES[0], Align=1,
        BackgroundColor=rgb(192, 192, 192),
        TextColor=rgb(255, 255, 255),
        FixedWidth=btnWidth, FixedHeight=btnHeight
    )]
    for day in LeenoUtils.DAYNAMES[1:]:
        dayNamesLabels.append(Dialogs.Spacer())
        dayNamesLabels.append(
            Dialogs.FixedText(Text=day, Align=1,
                BackgroundColor=workdaysBkColor if day not in ('Sab', 'Dom') else holydaysBkColor,
                TextColor=workdaysFgColor if day not in ('Sab', 'Dom') else holydaysFgColor,
                FixedWidth=btnWidth, FixedHeight=btnHeight
            )
        )

    def mkDayLabels():
        monthDay = 1
        weeks = []
        for week in range(0, 5):
            items = []
            id = str(week) + '.0'
            items.append(Dialogs.Button(
                Id=id, Label=str(monthDay),
                BackgroundColor=workdaysBkColor,
                TextColor=workdaysFgColor,
                FixedWidth=btnWidth, FixedHeight=btnHeight
            ))
            monthDay += 1
            for day in range(1, 7):
                items.append(Dialogs.Spacer())
                id = str(week) + '.' + str(day)
                items.append(Dialogs.Button(
                    Id=id, Label=str(monthDay),
                    BackgroundColor=workdaysBkColor if day not in (5, 6) else holydaysBkColor,
                    TextColor=workdaysFgColor if day not in (5, 6) else holydaysFgColor,
                    FixedWidth=btnWidth, FixedHeight=btnHeight
                ))
                monthDay += 1
            weeks.append(Dialogs.HSizer(Items=items))
            weeks.append(Dialogs.Spacer())
        return weeks

    def loadDate(dlg, dat):
        day = - LeenoUtils.firstWeekDay(dat) + 1
        days = LeenoUtils.daysInMonth(dat)
        for week in range(0, 5):
            for wDay in range(0, 7):
                id = str(week) + '.' + str(wDay)
                if day <= 0 or day > days:
                    dlg[id].setLabel('')
                else:
                    dlg[id].setLabel(str(day))
                day += 1

        dlg['date'].setText(LeenoUtils.date2String(dat))

    def handler(dlg, widgetId, widget, cmdStr):
        nonlocal curDate
        if widgetId == 'prevYear':
            curDate = date(year=curDate.year-1, month=curDate.month, day=curDate.day)
        elif widgetId == 'prevMonth':
            month = curDate.month - 1
            if month < 1:
                month = 12
                year = curDate.year - 1
            else:
                year = curDate.year
            curDate = date(year=year, month=month, day=curDate.day)
        elif widgetId == 'nextMonth':
            month = curDate.month + 1
            if month > 12:
                month = 1
                year = curDate.year + 1
            else:
                year = curDate.year
            curDate = date(year=year, month=month, day=curDate.day)
        elif widgetId == 'nextYear':
            curDate = date(year=curDate.year+1, month=curDate.month, day=curDate.day)
        elif widgetId == 'today':
            curDate = date.today()
        elif '.' in widgetId:
            txt = widget.getLabel()
            if txt == '':
                return
            day = int(txt)
            curDate = date(year=curDate.year, month=curDate.month, day=day)
        else:
            return
        loadDate(dlg, curDate)

    dlg = Dialogs.Dialog(Title='Selezionare la data', Horz=False, CanClose=True, Handler=handler, Items=[
        Dialogs.HSizer(Items=[
            Dialogs.Button(Id='prevYear', Icon='Icons-24x24/leftdbl.png'),
            Dialogs.Spacer(),
            Dialogs.Button(Id='prevMonth', Icon='Icons-24x24/leftsng.png'),
            Dialogs.Spacer(),
            Dialogs.FixedText(Id='date', Text='99 Settembre 9999', Align=1, FixedWidth=dateWidth),
            Dialogs.Spacer(),
            Dialogs.Button(Id='nextMonth', Icon='Icons-24x24/rightsng.png'),
            Dialogs.Spacer(),
            Dialogs.Button(Id='nextYear', Icon='Icons-24x24/rightdbl.png'),
        ]),
        Dialogs.Spacer(),
        Dialogs.HSizer(Items=dayNamesLabels),
        Dialogs.Spacer()
    ] + mkDayLabels() + [
        Dialogs.Spacer(),
        Dialogs.HSizer(Items=[
            Dialogs.Spacer(),
            Dialogs.Button(Label='Ok', MinWidth=MINBTNWIDTH, Icon='Icons-24x24/ok.png',  RetVal=1),
            Dialogs.Spacer(),
            Dialogs.Button(Id='today', Label='Oggi', MinWidth=MINBTNWIDTH),
            Dialogs.Spacer(),
            Dialogs.Button(Label='Annulla', MinWidth=MINBTNWIDTH, Icon='Icons-24x24/cancel.png',  RetVal=-1),
            Dialogs.Spacer()
        ])
    ])

    loadDate(dlg, curDate)
    if dlg.run() < 0:
        return None
    return curDate
