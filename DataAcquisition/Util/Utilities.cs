using DataAcquisition.Models;

namespace DataAcquisition.Util
{
    public static class Utilities
    {
        private static readonly Dictionary<DateTime?, decimal?>
            _currencyRateCache
                = new Dictionary<DateTime?, decimal?>();

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
            var a = context.Events
                .Where(x => x.Type == 6)
                .Select(x => new
                {
                    EventId = x.Id,
                    Rate = x.CurrencyPurchase.Price /
                           (decimal)x.CurrencyPurchase.Currency
                })
                .Average(x => x.Rate);
            return a;
        }
    }
}