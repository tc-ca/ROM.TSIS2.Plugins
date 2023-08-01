using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Plugins
{
    public class TimeZoneHelper
    {
        public static  DateTime GetAdjustedDateTime(ts_timezone timezone, DateTime sourceDateTime)
        {
            var timeZoneHoursAdjust = 0;
            var isDayLightSaving = 0;
            var timeZoneId = "Eastern Standard Time";
            TimeZoneInfo time_zone;

            switch (timezone)
            {
                case ts_timezone.AtlanticTime:
                    timeZoneHoursAdjust = -4;
                    timeZoneId = "Atlantic Standard Time";
                    break;
                case ts_timezone.CentralTime:
                    timeZoneHoursAdjust = -6;
                    timeZoneId = "Central Standard Time";
                    break;
                case ts_timezone.EasternTime:
                    timeZoneHoursAdjust = -5;
                    timeZoneId = "Eastern Standard Time";
                    break;
                case ts_timezone.MountainTime:
                    timeZoneHoursAdjust = -7;
                    timeZoneId = "Mountain Standard Time";
                    break;
                case ts_timezone.PacificTime:
                    timeZoneHoursAdjust = -8;
                    timeZoneId = "Pacific Standard Time";
                    break;
            }
            time_zone = TimeZoneInfo.FindSystemTimeZoneById(timeZoneId);

            if (time_zone.IsDaylightSavingTime(sourceDateTime))
            {
                isDayLightSaving = 1;
            }

            return sourceDateTime.AddHours(timeZoneHoursAdjust + isDayLightSaving);
        }
    }
}
