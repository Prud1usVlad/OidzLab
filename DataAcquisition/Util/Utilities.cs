using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataAcquisition.Models;

namespace DataAcquisition.Util
{
    public class Utilities
    {
        private static readonly Dictionary<DateTime?, decimal?> _currencyRateCache
            = new Dictionary<DateTime?, decimal?>();


        public static decimal? GetCurrencyRate(OidzDbContext context, DateTime? date)
        {
            if (_currencyRateCache.Keys.Contains(date))
                return _currencyRateCache[date];
            else
            {
                var data = context.Events.Where(e => e.Type == 6 && e.Date == date);

                decimal? rate = (decimal?)data.Sum(e => e.CurrencyPurchase.Price)
                    / (decimal?)data.Sum(e => e.CurrencyPurchase.Currency);

                _currencyRateCache[date] = rate;
                return rate;
            }
        }

    }
}
