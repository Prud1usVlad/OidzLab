using System.Text;
using DataAcquisition.Models;

namespace DataAcquisition.Util
{
    public static class Utilities
    {
        private static readonly Dictionary<DateTime?, decimal?>
            _currencyRateCache
                = new Dictionary<DateTime?, decimal?>();

        private static readonly decimal? _usdRate = null;

        private static List<int>? _ageGroups;

        private static readonly List<char> _alphabet = new List<char>();

        public static decimal? GetCurrencyRate(OidzDbContext context,
            DateTime? date)
        {
            if (_currencyRateCache.Keys.Contains(date))
                return _currencyRateCache[date];
            else
            {
                var data = context.Events.Where(e => e.Type == 6 &&
                                                     e.Date == date);
                decimal? rate = (decimal?)data.Sum(e =>
                                    e.CurrencyPurchase.Price)
                                / (decimal?)data.Sum(e =>
                                    e.CurrencyPurchase.Currency);
                _currencyRateCache[date] = rate;
                return rate;
            }
        }

        public static decimal? GetEventUSDRate(OidzDbContext context)
        {
            return _usdRate ?? context.Events
                .Where(x => x.Type == 6)
                .Select(x => new
                {
                    EventId = x.Id,
                    Rate = x.CurrencyPurchase.Price /
                           (decimal)x.CurrencyPurchase.Currency
                })
                .Average(x => x.Rate);
        }

        public static List<int> GetAgeGroups(OidzDbContext context, int groupsAmount = 5)
        {
            if (_ageGroups == null)
            {
                if (groupsAmount < 2)
                {
                    groupsAmount = 2;
                }
                
                var minAge = (int)context.Users.Min(x => x.Age);
                var maxAge = (int)context.Users.Max(x => x.Age) + 1;
                var diff = (int)Math.Ceiling((double)(maxAge - minAge) / groupsAmount);

                _ageGroups = new List<int>();
                
                var i = minAge;
                
                while (i < maxAge)
                {
                    _ageGroups.Add(i);
                    i += diff;
                }

                if (_ageGroups.Count < groupsAmount)
                {
                    _ageGroups.Add(maxAge);
                }
                else
                {
                    _ageGroups[^1] = maxAge;
                }
            }

            return _ageGroups;
        }

        public static string GetCellColumnAddress(int columnNumber)
        {
            if (!_alphabet.Any())
            {
                for (char c = 'a'; c <= 'z'; c++)
                {
                    _alphabet.Add(c);
                }
            }

            var cellColumnAddress = new StringBuilder(String.Empty);
            var _alphabetLenght = _alphabet.Count;

            if (columnNumber <= _alphabetLenght)
            {
                cellColumnAddress.Append(_alphabet[columnNumber - 1].ToString());
            }
            else
            {
                cellColumnAddress.Append(_alphabet[columnNumber / _alphabetLenght - 1].ToString());
                cellColumnAddress.Append(_alphabet[columnNumber % _alphabetLenght - 1].ToString());
            }
            
            return cellColumnAddress.ToString();
        }
    }
}