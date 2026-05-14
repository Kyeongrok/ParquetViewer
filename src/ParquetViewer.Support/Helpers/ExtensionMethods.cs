using Microsoft.Win32;
using ParquetViewer.Engine.Types;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Windows.Media.Imaging;

namespace ParquetViewer.Helpers
{
    public static class ExtensionMethods
    {
        private const string DefaultDateTimeFormat = "g";
        private const string DefaultDateOnlyFormat = "d";
        private const string DefaultTimeOnlyFormat = "T";
        public const string ISO8601DateTimeFormat = "yyyy-MM-ddTHH:mm:ss.FFFFFFF";
        public const string ISO8601DateOnlyFormat = "yyyy-MM-dd";
        public const string ISO8601TimeOnlyFormat = "HH:mm:ss.FFFFFFF";

        public static IList<string> GetColumnNames(this DataTable datatable)
        {
            List<string> columns = new List<string>(datatable.Columns.Count);
            foreach (System.Data.DataColumn column in datatable.Columns)
                columns.Add(column.ColumnName);
            return columns;
        }

        public static string GetDateFormat(this DateFormat dateFormat) => dateFormat switch
        {
            DateFormat.ISO8601 => ISO8601DateTimeFormat,
            DateFormat.Default => DefaultDateTimeFormat,
            DateFormat.Custom => AppSettings.CustomDateFormat ?? DefaultDateTimeFormat,
            _ => string.Empty
        };

        public static string GetDateOnlyFormat(this DateFormat dateFormat) => dateFormat switch
        {
            DateFormat.ISO8601 => ISO8601DateOnlyFormat,
            DateFormat.Default => DefaultDateOnlyFormat,
            DateFormat.Custom => AppSettings.CustomDateFormat is not null ?
                UtilityMethods.StripTimeComponentsFromDateTimeFormat(AppSettings.CustomDateFormat) : DefaultDateOnlyFormat,
            _ => string.Empty
        };

        public static string GetTimeOnlyFormat(this DateFormat dateFormat) => dateFormat switch
        {
            DateFormat.ISO8601 => ISO8601TimeOnlyFormat,
            DateFormat.Default => DefaultTimeOnlyFormat,
            DateFormat.Custom => AppSettings.CustomDateFormat is not null ?
                UtilityMethods.StripDateComponentsFromDateTimeFormat(AppSettings.CustomDateFormat) : DefaultTimeOnlyFormat,
            _ => string.Empty
        };

        public static string GetExtension(this FileType fileType)
            => Enum.IsDefined(fileType)
            ? $".{fileType.ToString().ToLowerInvariant()}"
            : throw new ArgumentOutOfRangeException(nameof(fileType));

        public static long ToMillisecondsSinceEpoch(this DateTime dateTime) => new DateTimeOffset(dateTime).ToUnixTimeMilliseconds();

        public static IEnumerable<DataColumn> AsEnumerable(this DataColumnCollection columns)
        {
            foreach (DataColumn column in columns)
                yield return column;
        }

        public static bool IsSimple(this Type type)
            => TypeDescriptor.GetConverter(type).CanConvertFrom(typeof(string));

        public static T ToEnum<T>(this int value, T @default) where T : struct, Enum
        {
            if (Enum.IsDefined(typeof(T), value))
                return (T)Enum.ToObject(typeof(T), value);
            return @default;
        }

        public static void DeleteSubKeyTreeIfExists(this RegistryKey key, string name)
        {
            if (key.OpenSubKey(name) is not null)
                key.DeleteSubKeyTree(name);
        }

        public static string? ToDecimalString(this float floatValue)
            => floatValue >= (float)decimal.MinValue && floatValue <= (float)decimal.MaxValue ? ToDecimalStringImpl(floatValue) : null;

        public static string? ToDecimalString(this double doubleValue)
            => doubleValue >= (double)decimal.MinValue && doubleValue <= (double)decimal.MaxValue ? ToDecimalStringImpl(doubleValue) : null;

        private static string? ToDecimalStringImpl(object value)
        {
            try { return Convert.ToDecimal(value).ToString(); }
            catch { return null; }
        }

        public static string Left(this string value, int maxLength, string? truncateSuffix = null)
        {
            if (string.IsNullOrEmpty(value)) return value;
            maxLength = Math.Abs(maxLength);
            return value.Length <= maxLength ? value : (value.Substring(0, maxLength) + truncateSuffix);
        }

        public static IEnumerable<T> AppendIf<T>(this IEnumerable<T> enumerable, bool append, T value)
        {
            if (append) return enumerable.Append(value);
            else return enumerable;
        }

        public static bool ToImage(this IByteArrayValue byteArrayValue, [NotNullWhen(true)] out BitmapImage? image)
        {
            ArgumentNullException.ThrowIfNull(byteArrayValue);
            try
            {
                using var ms = new MemoryStream(byteArrayValue.Data);
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.StreamSource = ms;
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
                image = bitmap;
                return true;
            }
            catch
            {
                image = null;
                return false;
            }
        }

        public static void DisposeSafely(this IDisposable? disposable)
        {
            try { disposable?.Dispose(); }
            catch { /*swallow*/ }
        }

        public static bool ImplementsInterface<T>(this Type? type)
        {
            if (type is null) return false;
            return typeof(T).IsAssignableFrom(type);
        }

        public static string Format(this string formatString, params object?[] args)
            => string.Format(formatString, args);

        public static bool IsHuggingFaceFormat(this IStructValue structValue, [NotNullWhen(true)] out byte[]? data)
        {
            if (structValue.Data.ColumnNames.Count == 2
                && structValue.Data.ColumnNames.Contains("bytes")
                && structValue.Data.ColumnNames.Contains("path")
                && structValue.Data.GetValue("bytes") is ByteArrayValue byteArrayValue)
            {
                data = byteArrayValue.Data;
                return true;
            }
            data = null;
            return false;
        }
    }
}
