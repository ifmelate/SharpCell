using SharpCell;

namespace SharpCell.Tests.Values;

public class DateSerialTests
{
    [Theory]
    [InlineData(0, 1900, 1, 0)]
    [InlineData(1, 1900, 1, 1)]
    [InlineData(59, 1900, 2, 28)]
    [InlineData(60, 1900, 2, 29)]
    [InlineData(61, 1900, 3, 1)]
    [InlineData(43831, 2020, 1, 1)]
    [InlineData(43831.75, 2020, 1, 1)]
    [InlineData(2958465, 9999, 12, 31)]
    public void Serial_to_date_in_1900(double serial, int year, int month, int day)
    {
        Assert.True(DateSerial.TryToDate(serial, DateSystem.Date1900, out var y, out var m, out var d));
        Assert.Equal((year, month, day), (y, m, d));
    }

    [Theory]
    [InlineData(0, 1904, 1, 1)]
    [InlineData(1, 1904, 1, 2)]
    [InlineData(42369, 2020, 1, 1)]
    public void Serial_to_date_in_1904(double serial, int year, int month, int day)
    {
        Assert.True(DateSerial.TryToDate(serial, DateSystem.Date1904, out var y, out var m, out var d));
        Assert.Equal((year, month, day), (y, m, d));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(2958466)]
    [InlineData(double.NaN)]
    public void Serial_outside_the_calendar_has_no_date(double serial)
    {
        Assert.False(DateSerial.TryToDate(serial, DateSystem.Date1900, out _, out _, out _));
    }

    [Theory]
    [InlineData(1900, 1, 1, 1)]
    [InlineData(1900, 2, 29, 60)]
    [InlineData(1900, 3, 0, 60)]
    [InlineData(1900, 3, 1, 61)]
    [InlineData(1900, 1, 0, 0)]
    [InlineData(2020, 1, 1, 43831)]
    [InlineData(2020, 13, 1, 44197)]
    [InlineData(2021, 0, 1, 44166)]
    [InlineData(2020, 2, 30, 43891)]
    [InlineData(2020, 1, -1, 43829)]
    [InlineData(120, 1, 1, 43831)]
    [InlineData(0, 1, 1, 1)]
    [InlineData(9999, 12, 31, 2958465)]
    public void Date_to_serial_like_DATE(long year, long month, long day, double expected)
    {
        Assert.True(DateSerial.TryFromDate(year, month, day, DateSystem.Date1900, out var serial));
        Assert.Equal(expected, serial);
    }

    [Theory]
    [InlineData(-1, 1, 1)]
    [InlineData(10000, 1, 1)]
    [InlineData(1900, 1, -1)]
    [InlineData(9999, 12, 32)]
    public void Date_outside_the_calendar_is_refused(long year, long month, long day)
    {
        Assert.False(DateSerial.TryFromDate(year, month, day, DateSystem.Date1900, out _));
    }

    [Fact]
    public void Date_to_serial_in_1904()
    {
        Assert.True(DateSerial.TryFromDate(1904, 1, 2, DateSystem.Date1904, out var serial));
        Assert.Equal(1, serial);
        Assert.False(DateSerial.TryFromDate(1903, 12, 31, DateSystem.Date1904, out _));
    }

    [Theory]
    [InlineData(0.5, 12, 0, 0)]
    [InlineData(43831.75, 18, 0, 0)]
    [InlineData(0.999988425925926, 23, 59, 59)]
    [InlineData(0.0000057870370370, 0, 0, 0)]
    [InlineData(0.0000057870370371, 0, 0, 1)]
    public void Time_of_day_rounds_to_the_second(double serial, int hour, int minute, int second)
    {
        Assert.Equal((hour, minute, second), DateSerial.TimeOfDay(serial));
    }

    [Fact]
    public void Round_trips_every_day()
    {
        for (var serial = 0; serial <= 2958465; serial += 7)
        {
            Assert.True(DateSerial.TryToDate(serial, DateSystem.Date1900, out var y, out var m, out var d));
            Assert.True(DateSerial.TryFromDate(y, m, d, DateSystem.Date1900, out var back));
            Assert.Equal(serial, back);
        }
    }

    [Theory]
    [InlineData(2020, 2, 30, false)]
    [InlineData(2020, 2, 29, true)]
    [InlineData(1900, 2, 29, true)]
    [InlineData(1899, 12, 31, false)]
    [InlineData(2020, 13, 1, false)]
    [InlineData(2020, 1, 0, false)]
    public void Calendar_dates_must_exist(int year, int month, int day, bool exists)
    {
        Assert.Equal(exists, DateSerial.TryFromCalendar(year, month, day, DateSystem.Date1900, out _));
    }

    [Fact]
    public void February_1900_has_29_days_only_in_the_1900_system()
    {
        Assert.Equal(29, DateSerial.DaysInMonth(1900, 2, DateSystem.Date1900));
        Assert.Equal(28, DateSerial.DaysInMonth(1900, 2, DateSystem.Date1904));
        Assert.False(DateSerial.TryFromCalendar(1900, 2, 29, DateSystem.Date1904, out _));
    }

    [Theory]
    [InlineData(1, DateSystem.Date1900, 6)]
    [InlineData(2, DateSystem.Date1900, 0)]
    [InlineData(0, DateSystem.Date1900, 5)]
    [InlineData(61, DateSystem.Date1900, 3)]
    [InlineData(43831.9, DateSystem.Date1900, 2)]
    [InlineData(0, DateSystem.Date1904, 4)]
    public void Weekdays_follow_the_serial(double serial, DateSystem system, int mondayBased)
    {
        Assert.Equal(mondayBased, DateSerial.MondayBasedWeekday(serial, system));
    }
}
