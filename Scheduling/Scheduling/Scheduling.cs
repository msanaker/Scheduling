using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Scheduling
{
    /// <summary>
    /// A set of Extension Methods for the .NET Framework System.DateTime structure
    /// </summary>
    /// <remarks>
    /// Copyright 2016 Matthew Sanaker, matthew@sanaker.com, @msanaker on GitHub
    ///
    /// This file is part of 'Scheduling'.
    ///
    /// 'Scheduling' is free software: you can redistribute it and/or modify
    /// it under the terms of the GNU General Public License as published by
    /// the Free Software Foundation, either version 3 of the License, or
    /// (at your option) any later version.
    ///
    /// 'Scheduling' is distributed in the hope that it will be useful,
    /// but WITHOUT ANY WARRANTY; without even the implied warranty of
    /// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
    /// GNU General Public License for more details.
    ///
    /// You should have received a copy of the GNU General Public License
    /// along with Scheduling.  If not, see<http://www.gnu.org/licenses/>.
    /// </remarks>
    public static class Scheduling
    {
        /// <summary>
        /// Find the number of workdays between two dates with the option of including Saturday
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="workOnSaturday"></param>
        /// <returns></returns>
        public static int FindNumberOfWorkdays(DateTime startDate, DateTime endDate, bool workOnSaturday, bool accountForHolidaysInUS)
        {
            try
            {
                DateTime beginningDate = startDate;
                DateTime endingDate = endDate;
                bool endBeforeStart = false;
                if (endDate < startDate)
                {
                    endBeforeStart = true;
                    beginningDate = endDate;
                    endingDate = startDate;
                }
                int weeks = 0;
                int workDaysInFinalWeek = 0;
                int daysBetween =(endingDate - beginningDate).Days;
                int workDaysInAWeek = 5;
                if (workOnSaturday) workDaysInAWeek = 6;
                int workDaysInFirstWeek = daysBetween;
                int daysToWeekend = workDaysInAWeek - ((int)beginningDate.DayOfWeek);
                if(daysBetween > daysToWeekend)
                {
                    int daysRemaining = 0;
                    workDaysInFirstWeek = workDaysInAWeek - ((int)beginningDate.DayOfWeek - 1);
                    daysRemaining = daysBetween - workDaysInFirstWeek;
                    decimal totalWeeks = daysRemaining / 7;
                    weeks = (int)Math.Floor(totalWeeks);
                    int daysInFinalWeek = daysRemaining % 7;
                    workDaysInFinalWeek = 0;
                    if (daysInFinalWeek > 1) workDaysInFinalWeek = daysInFinalWeek - (7 - workDaysInAWeek);
                }
                int TotalWorkDays = workDaysInFirstWeek + (weeks * workDaysInAWeek) + workDaysInFinalWeek;
                HashSet<DateTime> holidays = new HashSet<DateTime>();
                if (beginningDate.Year == endDate.Year)
                {
                    holidays = getHolidaysForUS(beginningDate.Year);
                }
                else
                {
                    holidays = getHolidaysForUS(beginningDate.Year);
                    HashSet<DateTime> holidays2 = getHolidaysForUS(endDate.Year);
                    foreach (DateTime date in holidays2)
                    {
                        holidays.Add(date);
                    }
                }
                foreach (DateTime date in holidays)
                {
                    if (date > beginningDate.Date && date < endingDate.Date) TotalWorkDays--;
                }
                if (endBeforeStart) TotalWorkDays = TotalWorkDays * -1;
                return TotalWorkDays;
            }
            catch (Exception ex)
            {
                Exception except = new Exception(string.Format("Exception thrown in {0}.{1} -- Inner Exception:\n  {2}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name, ex.Message), ex);
                throw except;
            }
        }

        /// <summary>
        /// DateTime extension method to add the given number of workdays, i.e. skip weekends, with the option of including Saturday as a workday.  Adding negative days will subtract them from the given date.
        /// </summary>
        /// <param name="assignedDate"></param>
        /// <param name="daysToAdd"></param>
        /// <param name="workOnSaturday"></param>
        /// <returns></returns>
        public static DateTime AddWorkDays(this DateTime assignedDate, int daysToAdd, bool workOnSaturday, bool accountForHolidaysInUS)
        {
            try
            {
                if (daysToAdd == 0) return assignedDate;
                int workDaysInAWeek = 5;
                if (workOnSaturday) workDaysInAWeek = 6;

                DateTime newDate = assignedDate;

                if(daysToAdd > 0)
                {
                    int workDaysInFirstWeek = workDaysInAWeek - ((int)assignedDate.DayOfWeek - 1);
                    if (assignedDate.DayOfWeek == DayOfWeek.Sunday) workDaysInFirstWeek = workDaysInAWeek;

                    if (daysToAdd > workDaysInFirstWeek)
                    {
                        int daysRemaining = daysToAdd - workDaysInFirstWeek;
                        int weeks = 0;
                        if(daysRemaining > workDaysInAWeek)
                        {
                            decimal addedWeeks = daysRemaining / workDaysInAWeek;
                            weeks = (int)Math.Floor(addedWeeks);
                        }
                        int daysInFinalWeek = daysRemaining % workDaysInAWeek;
                        int calendarDaysToAdd = workDaysInFirstWeek + (weeks * 7);
                        newDate = assignedDate.AddDays(calendarDaysToAdd);
                        if (newDate.DayOfWeek == DayOfWeek.Saturday && (!(workOnSaturday))) newDate = newDate.AddDays(2);
                        if (newDate.DayOfWeek == DayOfWeek.Sunday) newDate = newDate.AddDays(1);
                        newDate = newDate.AddDays(daysInFinalWeek);
                    }
                    else
                    {
                        newDate = assignedDate.AddDays(daysToAdd);
                        if (newDate.DayOfWeek == DayOfWeek.Saturday && (!(workOnSaturday))) newDate = newDate.AddDays(2);
                        if (newDate.DayOfWeek == DayOfWeek.Sunday) newDate = newDate.AddDays(1);
                    }
                }
                else
                {
                    int workDaysInFirstWeek = -1 * (int)assignedDate.DayOfWeek;
                    if (workOnSaturday) newDate = assignedDate.AddDays(-1);

                    if (daysToAdd < workDaysInFirstWeek)
                    {
                        int daysRemaining = daysToAdd - workDaysInFirstWeek;
                        int weeks = 0;
                        if (daysRemaining < workDaysInAWeek)
                        {
                            decimal addedWeeks = daysRemaining / workDaysInAWeek;
                            weeks = (int)Math.Floor(addedWeeks);
                        }
                        int daysInFinalWeek = daysRemaining % workDaysInAWeek;
                        int calendarDaysToAdd = workDaysInFirstWeek + (weeks * 7) + daysInFinalWeek;
                        newDate = newDate.AddDays(calendarDaysToAdd);
                        if (newDate.DayOfWeek == DayOfWeek.Saturday && (!(workOnSaturday))) newDate = newDate.AddDays(-1);
                        if (newDate.DayOfWeek == DayOfWeek.Sunday) newDate = newDate.AddDays(-2);
                    }
                    else
                    {
                        newDate = newDate.AddDays(daysToAdd);
                        if (newDate.DayOfWeek == DayOfWeek.Saturday && (!(workOnSaturday))) newDate = newDate.AddDays(-1);
                        if (newDate.DayOfWeek == DayOfWeek.Sunday) newDate = newDate.AddDays(-2);
                    }
                }
                return newDate;
            }
            catch (Exception ex)
            {
                Exception except = new Exception(string.Format("Exception thrown in {0}.{1} -- Inner Exception:\n  {2}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name, ex.Message), ex);
                throw except;
            }
        }

        /// <summary>
        /// DateTime extension method to adjust a due date to the next workday with the option to account for weekends, US holidays and treat Saturdays as a work day.
        /// </summary>
        /// <param name="dueDate"></param>
        /// <param name="accountForWeekends"></param>
        /// <param name="accountForHolidaysInUS"></param>
        /// <param name="workOnSaturday"></param>
        /// <returns></returns>
        public static DateTime adjustDateToNextWorkDay(this DateTime dueDate, bool accountForWeekends, bool accountForHolidaysInUS, bool workOnSaturday)
        {
            try
            {
                HashSet<DateTime> holidays = getHolidaysForUS(dueDate.Year);
                DateTime newDate = dueDate;
                while(accountForWeekends && (newDate.DayOfWeek == DayOfWeek.Saturday || newDate.DayOfWeek == DayOfWeek.Sunday))
                {
                    newDate = moveWeekendDateToMonday(newDate, workOnSaturday);
                    while (accountForHolidaysInUS && holidays.Any(d => d.Date == newDate.Date))
                    {
                        newDate = newDate.AddDays(1);
                    }
                }

                while (accountForHolidaysInUS && holidays.Any(d => d.Date == newDate.Date))
                {
                    newDate = newDate.AddDays(1);
                    while ((accountForWeekends && (newDate.DayOfWeek == DayOfWeek.Saturday || newDate.DayOfWeek == DayOfWeek.Sunday)))
                    {
                        newDate = moveWeekendDateToMonday(newDate, workOnSaturday);
                    }
                }

                return newDate;
            }
            catch (Exception ex)
            {
                Exception except = new Exception(string.Format("Exception thrown in {0}.{1} -- Inner Exception:\n  {2}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name, ex.Message), ex);
                throw except;
            }
        }

        /// <summary>
        /// Used internally to find the Monday following the weekend date passed as the date parameter
        /// </summary>
        /// <param name="date"></param>
        /// <param name="workOnSaturday"></param>
        /// <returns></returns>
        private static DateTime moveWeekendDateToMonday(DateTime date, bool workOnSaturday)
        {
            try
            {
                DateTime newDate = date;
                if (!workOnSaturday && newDate.DayOfWeek == DayOfWeek.Saturday) newDate = newDate.AddDays(1);
                if (newDate.DayOfWeek == DayOfWeek.Sunday) newDate = newDate.AddDays(1);
                return newDate;
            }
            catch (Exception ex)
            {
                Exception except = new Exception(string.Format("Exception thrown in {0}.{1} -- Inner Exception:\n  {2}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name, ex.Message), ex);
                throw except;
            }
        }

        /// <summary>
        /// Return a HashSet of DateTime objects that are the standard holidays for the supplied year including New Year's day of the following year.
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        /// <remarks>
        /// The initial code template that I used for this method was written by Wil Peck and can be found here:
        /// http://geekswithblogs.net/wpeck/archive/2011/12/27/us-holiday-list-in-c.aspx
        /// </remarks>
        public static HashSet<DateTime> getHolidaysForUS(int year)
        {
            try
            {
                HashSet<DateTime> holidays = new HashSet<DateTime>();
                // New Year's
                holidays.Add(AdjustForWeekendHoliday(new DateTime(year, 1, 1).Date)); 
                // Memorial Day  -- last monday in May
                DateTime memorialDay = new DateTime(year, 5, 31);
                DayOfWeek dayOfWeek = memorialDay.DayOfWeek;
                while (dayOfWeek != DayOfWeek.Monday)
                {
                    memorialDay = memorialDay.AddDays(-1);
                    dayOfWeek = memorialDay.DayOfWeek;
                }
                holidays.Add(memorialDay.Date);
                //Independence Day
                holidays.Add(AdjustForWeekendHoliday(new DateTime(year, 7, 4).Date));
                //Labor Day -- 1st Monday in September
                DateTime laborDay = new DateTime(year, 9, 1);
                dayOfWeek = laborDay.DayOfWeek;
                while (dayOfWeek != DayOfWeek.Monday)
                {
                    laborDay = laborDay.AddDays(1);
                    dayOfWeek = laborDay.DayOfWeek;
                }
                holidays.Add(laborDay.Date);
                //Thanksgiving - 4th Thursday in November
                var thanksgiving = (from day in Enumerable.Range(1, 30)
                                    where new DateTime(year, 11, day).DayOfWeek == DayOfWeek.Thursday
                                    select day).ElementAt(3);
                DateTime thanksgivingDay = new DateTime(year, 11, thanksgiving);
                holidays.Add(thanksgivingDay.Date);
                //Christmas Eve & Day
                holidays.Add(AdjustForWeekendHoliday(new DateTime(year, 12, 24).Date));
                holidays.Add(AdjustForWeekendHoliday(new DateTime(year, 12, 25).Date));
                //New Year's Eve Day & New Year of following year
                holidays.Add(AdjustForWeekendHoliday(new DateTime(year, 12, 31).Date));
                year++;
                holidays.Add(AdjustForWeekendHoliday(new DateTime(year, 1, 1).Date));
                return holidays;
            }
            catch (Exception ex)
            {
                Exception except = new Exception(string.Format("Exception thrown in {0}.{1} -- Inner Exception:\n  {2}", MethodBase.GetCurrentMethod().DeclaringType.FullName, MethodBase.GetCurrentMethod().Name, ex.Message), ex);
                throw except;
            }
        }

        /// <summary>
        /// Used internally to move a holiday to the Friday prior or Monday after a holiday that falls on a weekend
        /// </summary>
        /// <param name="holiday"></param>
        /// <returns></returns>
        /// <remarks>
        /// The initial code template that I used for this method was written by Wil Peck and can be found here:
        /// http://geekswithblogs.net/wpeck/archive/2011/12/27/us-holiday-list-in-c.aspx
        /// </remarks>
        private static DateTime AdjustForWeekendHoliday(DateTime holiday)
        {
            if (holiday.DayOfWeek == DayOfWeek.Saturday)
            {
                return holiday.AddDays(-1);
            }
            else if (holiday.DayOfWeek == DayOfWeek.Sunday)
            {
                return holiday.AddDays(1);
            }
            else
            {
                return holiday;
            }
        }

    }
}
