using System;
namespace WebView.Models
{
    interface ICabInfoModel
    {
        string AIRCOND_BROKE { get; set; }
        string AIRCOND_DOOR { get; set; }
        string AIRCOND_READING { get; set; }
        string AIRCOND_WORK { get; set; }
        string ALARM { get; set; }
        string BATT_BRAND { get; set; }
        string BATT_CAPASITY { get; set; }
        string BATT_CONDITION { get; set; }
        string BATT_DATE { get; set; }
        string BATT_VOLT { get; set; }
        string CHECK_BY { get; set; }
        string EQUIP_CAB_CONDITION { get; set; }
        string EXC_ABB { get; set; }
        string G3E_FID { get; set; }
        string PRO_AC_CONNECTION { get; set; }
        string PRO_AC_OHM { get; set; }
        string PRO_ELCB_AUTO_BRAND { get; set; }
        string PRO_ELCB_CAPASITY { get; set; }
        string PRO_ELCB_DATE { get; set; }
        string PRO_ELCB_LOAC { get; set; }
        string PRO_ELCB_VAC { get; set; }
        string PRO_MCB_BRAND { get; set; }
        string PRO_MCB_CAPASITY { get; set; }
        string PRO_MCB_DATE { get; set; }
        string PRO_SA_BRAND { get; set; }
        string PRO_SA_CAPASITY { get; set; }
        string PRO_SA_CONDITION { get; set; }
        string PTT_ID { get; set; }
        string PWR_CAB_CONDITION { get; set; }
        string PWR_CAB_ID { get; set; }
        string RC_BMODULES { get; set; }
        string RC_BRAND { get; set; }
        string RC_DATE { get; set; }
        string RC_GMODULES { get; set; }
        string RC_LVD { get; set; }
        string RC_MCB_BRAND { get; set; }
        string RC_MCB_CAPASITY { get; set; }
        string RC_MODEL { get; set; }
        string RC_PRESENT { get; set; }
        string RC_RATING { get; set; }
        string RC_SA_BRAND { get; set; }
        string RC_SA_CAPASITY { get; set; }
        string RC_SA_CONDITION { get; set; }
        string RC_SA_CONNECTION { get; set; }
        string RC_SA_OHM { get; set; }
        string RC_VAC { get; set; }
        string RC_VDC { get; set; }
    }
}
