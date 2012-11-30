using Sitecore;

namespace ASR.Controls
{
    public class DateTimePicker : Sitecore.Web.UI.HtmlControls.DateTimePicker
    {
        public string Format
        {
            get { return base.GetViewStateString("Value.Format"); }
            set
            {
                base.SetViewStateString("Value.Format", value);
            }
        }

        public DateTimePicker():base()
        {
            Format = "yyyyMMddTHHmmss";
        }

        protected override void DoRender(System.Web.UI.HtmlTextWriter output)
        {
            base.DoRender(output);
        }
        public override string Value
        {
            get
            {
                var value = base.Value;
                if(DateUtil.IsIsoDate(value))
                {                    
                    return DateUtil.IsoDateToDateTime(value).ToString(Format);
                }
                return value;
            }
            set
            {                
               base.Value = Util.MakeDateReplacements(value);
            }
        }
    }
}
