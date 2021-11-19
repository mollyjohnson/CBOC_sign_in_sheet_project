import calendar

####################################################################
### Function Title: calcFedHolidays()
### Arguments: datetime object
### Returns: list of federal holiday dates for a given month
### Description: calculates all federal holiday dates for a given month
### (month provided by the datetime object). will adjust for if the
### federal holiday occurs on a saturday or sunday as needed
####################################################################    
def calcFedHolidays(dateTimeObj):
    holidayDates = []
    curDate = 1
    endDate = calendar.monthrange(dateTimeObj.year, dateTimeObj.month)[1]
    monthName = calendar.month_name[dateTimeObj.month]
    monthName = monthName.upper()
    dayDate = dateTimeObj.weekday()
    dayName = calendar.day_abbr[dayDate]
    dayName = dayName.upper()
    mlkMondays = 0
    washBdayMondays = 0
    memorialDayLastMonInMayDate = 0
    laborDayMon = 0
    columbusDayMon = 0
    thanksgivingThurs = 0

    # loop through every day in the month and add holiday dates to the list as appropriate
    while(curDate <= endDate):
        if(monthName == "JANUARY"):
            if(curDate == 1):
                # if new year's occurs on a sat, will be taken care of in dec
                # if new year's occurs on a sun, push it to mon
                if(dayName == "SUN"):
                    holidayDates.append(curDate + 1)
                elif(dayName != "SAT"):
                    holidayDates.append(curDate)
            if(dayName == "MON"):
                mlkMondays += 1
                # mlk bday falls on 3rd monday of january
                if(mlkMondays == 3):
                    holidayDates.append(curDate)
        if(monthName == "FEBRUARY"):
            if(dayName == "MON"):
                washBdayMondays += 1
                # washington's bday falls on 3rd monday of the month
                if(washBdayMondays == 3):
                    holidayDates.append(curDate)
        # no fed hols for march or april
        if(monthName == "MAY"):
            # keep replacing last monday in may w the current one,
            # such that the last monday in may will be the last value
            # assigned to the variable (which can then be added later)
            if(dayName == "MON"):
                memorialDayLastMonInMayDate = curDate
        if(monthName == "JUNE"):
            # juneteenth holiday
            if(curDate == 19):
                # if 19th occurs on a sat, add fri to list
                if(dayName == "SAT"):
                    holidayDates.append(curDate - 1)
                # if 19th holiday occurs on a sun, add mon to list
                elif(dayName == "SUN"):
                    holidayDates.append(curDate + 1)
                else:
                    holidayDates.append(curDate)
        if(monthName == "JULY"):
            # 4th of july holiday
            if(curDate == 4):
                # if 4th holiday occurs on a sat, add fri to list.
                if(dayName == "SAT"):
                    holidayDates.append(curDate - 1)
                # if 4th holiday occurs on a sun, add mon to list
                elif(dayName == "SUN"):
                    holidayDates.append(curDate + 1)
                else:
                    holidayDates.append(curDate)
        # no fed hols for august
        if(monthName == "SEPTEMBER"):
            if(dayName == "MON"):
                laborDayMon += 1
                # if is first mon of the month, that's labor day
                if(laborDayMon == 1):
                    holidayDates.append(curDate)
        if(monthName == "OCTOBER"):
            if(dayName == "MON"):
                columbusDayMon += 1
                # if is second mon of month, that's columbus day
                if(columbusDayMon == 2):
                    holidayDates.append(curDate)
        if(monthName == "NOVEMBER"):
            # veteran's day celebrated 11th of nov
            if(curDate == 11):
                # if 11h is a sat, add fri to hols list
                if(dayName == "SAT"):
                    holidayDates.append(curDate - 1)
                # if 11th is a sun, add mon to hols list
                elif(dayName == "SUN"):
                    holidayDates.append(curDate + 1)
                else:
                    holidayDates.append(curDate)
            if(dayName == "THU"):
                thanksgivingThurs += 1
                # if is 4th thurs of month, that's thanksgiving
                if(thanksgivingThurs == 4):
                    holidayDates.append(curDate)
        if(monthName == "DECEMBER"):
            # if new year's occurs on a sat, add fri dec 31 to hols
            if((curDate == 31) and (dayName == "FRI")):
                holidayDates.append(curDate) 
            # 25th = christmas
            if(curDate == 25):
                # if 25th is on a sat, add fri to hols list
                if(dayName == "SAT"):
                    holidayDates.append(curDate - 1)
                # if 25th is on a sun, add mon to hols list
                elif(dayName == "SUN"):
                    holidayDates.append(curDate + 1)
                else:
                    holidayDates.append(curDate)

        curDate += 1
        dayDate += 1
        #if day number is > 6, i.e. you've reached end of week, start week days over
        if (dayDate > 6):
            dayDate = 0
        dayName = calendar.day_abbr[dayDate]
        dayName = dayName.upper()

    # if month is may, add last monday for memorial day
    if(monthName == "MAY"):
        holidayDates.append(memorialDayLastMonInMayDate)

    return holidayDates

