using Sowfin.API.ViewModels.PayoutPolicy;
using Sowfin.Data.Common.Helper;
public class CurrentSetuoSoResult
{
    public int StatusCode { get; set; }
    public string Message { get; set; }
    public List<CurrentSetupSoFillingsViewModel> Result { get; set; }
    public List<ValueTextWrapper> currencyValueList { get; set; } //Date:23-July-2020 | Added By : anonymous | Enh. : Single Units Enum 
    public List<ValueTextWrapper> numberCountList { get; set; } //Date:23-July-2020 | Added By : anonymous | Enh. : Single Units Enum 
}