using Sowfin.Data.Common.Enum;
using Sowfin.Data.Common.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sowfin.Data.Repositories
{
    public class EnumListRepositoryManager
    {

      

        public List<FAnalysisFlag> getFAnalysisFlagList()
        {
            List<FAnalysisFlag> FlagList = new List<FAnalysisFlag>();
            try
            {
                Array values = Enum.GetValues(typeof(FAnalysisFlagEnum));
                for (int i = 1; i <= values.Length; i++)
                {
                    FlagList.Add(new FAnalysisFlag
                    {
                        FlagName = EnumHelper.DescriptionAttr((FAnalysisFlagEnum)i),
                        Id = i,
                        FlagValue = false
                    });
                }
                FlagList = FlagList.OrderBy(x => x.Id).ToList();
            }
            catch (Exception ss)
            {

            }
            return FlagList;
        }

        //get Value Type Enum list
        public List<SelectListItem> GetValueTypeEnumListforDropdown()
        {
            List<SelectListItem> itemsList = new List<SelectListItem>();
            try
            {
                Array values = Enum.GetValues(typeof(ValueTypeEnum));
                for (int i = 1; i <= values.Length; i++)
                {
                    itemsList.Add(new SelectListItem
                    {
                        Text = EnumHelper.DescriptionAttr((ValueTypeEnum)i),
                        Value = i

                    });
                }
                itemsList = itemsList.OrderBy(x => x.Value).ToList();
            }
            catch (Exception ss)
            {

            }
            return itemsList;
        }

        //get Leverage Policy Enum list
        public List<SelectListItem> GetLeveragePolicyEnumListforDropdown()
        {
            List<SelectListItem> itemsList = new List<SelectListItem>();
            try
            {
                Array values = Enum.GetValues(typeof(LeveragePolicyEnum));
                for (int i = 1; i <= values.Length; i++)
                {
                    itemsList.Add(new SelectListItem
                    {
                        Text = EnumHelper.DescriptionAttr((LeveragePolicyEnum)i),
                        Value = i

                    });
                }
                itemsList = itemsList.OrderBy(x => x.Value).ToList();
            }
            catch (Exception ss)
            {

            }
            return itemsList;
        }


        //get Header Enum list
        public List<SelectListItem> GetHeaderEnumListforDropdown()
        {
            List<SelectListItem> itemsList = new List<SelectListItem>();
            try
            {
                Array values = Enum.GetValues(typeof(HeadersEnum));
                for (int i = 1; i <= values.Length; i++)
                {
                    itemsList.Add(new SelectListItem
                    {
                        Text = EnumHelper.DescriptionAttr((HeadersEnum)i),
                        Value = i

                    });
                }
                itemsList = itemsList.OrderBy(x => x.Value).ToList();
            }
            catch (Exception ss)
            {

            }
            return itemsList;
        }


        //get Beta Source Enum list
        public List<SelectListItem> GetBetaSourceEnumListforDropdown()
        {
            List<SelectListItem> itemsList = new List<SelectListItem>();
            try
            {
                Array values = Enum.GetValues(typeof(BetaSourceEnum));
                for (int i = 1; i <= values.Length; i++)
                {
                    itemsList.Add(new SelectListItem
                    {
                        Text = EnumHelper.DescriptionAttr((BetaSourceEnum)i),
                        Value = i

                    });
                }
                itemsList = itemsList.OrderBy(x => x.Value).ToList();
            }
            catch (Exception ss)
            {

            }
            return itemsList;
        }


        //get NUmber count Enum list
        public List<SelectListItem> GetNumberCountEnumListforDropdown()
        {
            List<SelectListItem> itemsList = new List<SelectListItem>();
            try
            {
                Array values = Enum.GetValues(typeof(NumberCountEnum));
                for (int i = 1; i <= values.Length; i++)
                {
                    itemsList.Add(new SelectListItem
                    {
                        Text = EnumHelper.DescriptionAttr((NumberCountEnum)i),
                        Value = i
                       
                    });
                }
                itemsList = itemsList.OrderBy(x => x.Value).ToList();
            }
            catch (Exception ss)
            {
                
            }
            return itemsList;
        }

        //get Currency Value Enum list
        public List<SelectListItem> GetCurrencyValueEnumListforDropdown()
        {
            List<SelectListItem> itemsList = new List<SelectListItem>();
            try
            {
                Array values = Enum.GetValues(typeof(CurrencyValueEnum));
                for (int i = 1; i <= values.Length; i++)
                {
                    itemsList.Add(new SelectListItem
                    {
                        Text = EnumHelper.DescriptionAttr((CurrencyValueEnum)i),
                        Value = i

                    });
                }
                itemsList = itemsList.OrderBy(x => x.Value).ToList();
            }
            catch (Exception ss)
            {

            }
            return itemsList;
        }
        
    }

    public class SelectListItem
    {
        public int Value { get; set; }
        public string Text { get; set; }
    }

    public class FAnalysisFlag
    {
        public int Id { get; set; }
        public string FlagName { get; set; }
        public bool FlagValue { get; set; }
    }
}
