using System;
using System.ComponentModel.DataAnnotations;

namespace WebView.Models
{
    public class CabInfoModel 
    {
        [Display(Name = "REGION")]
        public string REGION { get; set; }

        [Display(Name = "STATE")]
        public string STATE { get; set; }

        [Display(Name = "SEGMENT")]
        public string SEGMENT { get; set; }

        [Display(Name = "ZONE")]
        public string ZONE { get; set; }

        [Display(Name = "PTT_ID")]
        public string PTT_ID { get; set; }

        [Display(Name = "EXC_ABB")]
        public string EXC_ABB { get; set; }

        [Display(Name = "G3E_FID")]
        public string G3E_FID { get; set; }

        [Display(Name = "G3E_FID")]
        public string G3E_FID2 { get; set; }

        [Required]
        [Display(Name = "PWR_CAB_ID")]
        public string PWR_CAB_ID { get; set; }

        [Required]
        [Display(Name = "PWR_CAB_CONDITION")]
        public string PWR_CAB_CONDITION { get; set; }

        [Required]
        [Display(Name = "EQUIP_CAB_CONDITION")]
        public string EQUIP_CAB_CONDITION { get; set; }

        [Display(Name = "PRO_ELCB_AUTO_BRAND")]
        public string PRO_ELCB_AUTO_BRAND { get; set; }

        [Display(Name = "PRO_ELCB_BRAND")]
        public string PRO_ELCB_BRAND { get; set; }

        [Display(Name = "PRO_ELCB_VAC")]
        public string PRO_ELCB_VAC { get; set; }

        [Display(Name = "PRO_ELCB_LOAC")]
        public string PRO_ELCB_LOAC { get; set; }

        [Display(Name = "PRO_ELCB_CAPASITY")]
        public string PRO_ELCB_CAPASITY { get; set; }

        DateTime DPRO_ELCB_DATE = DateTime.Now;//set Default Value here
        [Display(Name = "PRO_ELCB_DATE")]
        [DataType(DataType.Date)]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMMM-yyyy}")]
        public DateTime? PRO_ELCB_DATE { get; set; }
        //{
        //    get
        //    {
        //        return DPRO_ELCB_DATE;
        //    }
        //    set
        //    {
        //        if (PRO_ELCB_DATE != null)
        //        {
        //            DPRO_ELCB_DATE = value;// (System.DateTime)PRO_ELCB_DATE;

        //        }
        //    }
        //}

        [Display(Name = "PRO_MCB_BRAND")]
        public string PRO_MCB_BRAND { get; set; }

        [Display(Name = "PRO_MCB_CAPASITY")]
        public string PRO_MCB_CAPASITY { get; set; }

        DateTime DPRO_MCB_DATE = DateTime.Now;//set Default Value here
        [Display(Name = "PRO_MCB_DATE")]
        [DataType(DataType.Date)]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMMM-yyyy}")]
        public DateTime? PRO_MCB_DATE { get; set; }
        //{
        //    get
        //    {
        //        return DPRO_MCB_DATE;
        //    }
        //    set
        //    {
        //        if (PRO_MCB_DATE != null)
        //        {
        //            DPRO_MCB_DATE = value;//(System.DateTime)PRO_MCB_DATE;

        //        }
        //    }
        //}

        [Display(Name = "PRO_SA_BRAND")]
        public string PRO_SA_BRAND { get; set; }

        [Display(Name = "PRO_SA_CAPASITY")]
        public string PRO_SA_CAPASITY { get; set; }

        [Display(Name = "PRO_SA_CONDITION")]
        public string PRO_SA_CONDITION { get; set; }

        [Display(Name = "PRO_AC_OHM")]
        public string PRO_AC_OHM { get; set; }

        [Display(Name = "PRO_AC_CONNECTION ")]
        public string PRO_AC_CONNECTION { get; set; }

        [Display(Name = "RC_BRAND")]
        public string RC_BRAND { get; set; }

        [Display(Name = "RC_MODEL")]
        public string RC_MODEL { get; set; }

        [Display(Name = "RC_RATING")]
        public string RC_RATING { get; set; }

        [Display(Name = "RC_GMODULES")]
        public string RC_GMODULES { get; set; }

        [Display(Name = "RC_BMODULES")]
        public string RC_BMODULES { get; set; }

        DateTime DRC_DATE = DateTime.Now;//set Default Value here
        [Display(Name = "RC_DATE")]
        [DataType(DataType.Date)]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMMM-yyyy}")]
        public DateTime? RC_DATE { get; set; }
        //{
        //    get
        //    {
        //        return DRC_DATE;
        //    }
        //    set
        //    {
        //        if (RC_DATE != null)
        //        {
        //            DRC_DATE = value;//(System.DateTime)RC_DATE;

        //        }
        //    }
        //}

        [Display(Name = "string RC_VAC")]
        public string RC_VAC { get; set; }

        [Display(Name = "RC_VDC")]
        public string RC_VDC { get; set; }

        [Display(Name = "RC_PRESENT")]
        public string RC_PRESENT { get; set; }

        [Display(Name = "RC_MCB_BRAND")]
        public string RC_MCB_BRAND { get; set; }

        [Display(Name = "RC_MCB_CAPASITY")]
        public string RC_MCB_CAPASITY { get; set; }

        [Display(Name = "RC_SA_BRAND")]
        public string RC_SA_BRAND { get; set; }

        [Display(Name = "RC_SA_CAPASITY")]
        public string RC_SA_CAPASITY { get; set; }

        [Display(Name = "RC_SA_CONDITION")]
        public string RC_SA_CONDITION { get; set; }

        [Display(Name = "RC_SA_OHM")]
        public string RC_SA_OHM { get; set; }

        [Display(Name = "RC_SA_CONNECTION")]
        public string RC_SA_CONNECTION { get; set; }

        [Display(Name = "RC_LVD")]
        public string RC_LVD { get; set; }

        [Display(Name = "BATT_BRAND")]
        public string BATT_BRAND { get; set; }

        [Display(Name = "BATT_CAPASITY")]
        public string BATT_CAPASITY { get; set; }

        [Display(Name = "BATT_VOLT")]
        public string BATT_VOLT { get; set; }

        DateTime DBATT_DATE = DateTime.Now;//set Default Value here
        [Display(Name = "BATT_DATE")]
        [DataType(DataType.Date)]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMMM-yyyy}")]
        public DateTime? BATT_DATE { get; set; }
        //{
        //    get
        //    {
        //        return DBATT_DATE;
        //    }
        //    set
        //    {
        //        if (DBATT_DATE != null)
        //        {
        //            DBATT_DATE = value;//(System.DateTime)DBATT_DATE;

        //        }
        //    }
        //}

        [Display(Name = "BATT_CONDITION")]
        public string BATT_CONDITION { get; set; }

        [Display(Name = "AIRCOND_READING")]
        public string AIRCOND_READING { get; set; }

        [Display(Name = "AIRCOND_WORK")]
        public string AIRCOND_WORK { get; set; }

        [Display(Name = "AIRCOND_BROKE")]
        public string AIRCOND_BROKE { get; set; }

        [Display(Name = "AIRCOND_DOOR")]
        public string AIRCOND_DOOR { get; set; }

        [Display(Name = "ALARM")]
        public string ALARM { get; set; }

        [Display(Name = "CHECK_BY")]
        public string CHECK_BY { get; set; }

        [Display(Name = "GROUP_ID")]
        public string GROUP_ID { get; set; }

        [Display(Name = "MANUFACTURER")]
        public string MANUFACTURER { get; set; }

        [Display(Name = "CONTRACTOR")]
        public string CONTRACTOR { get; set; }

        public string TNB_METER { get; set; }

        DateTime CCAB_INS_DATE = DateTime.Now;//set Default Value here
        [Display(Name = "CAB_INS_DATE")]
        [DataType(DataType.Date)]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMMM-yyyy}")]
        public DateTime? CAB_INS_DATE { get; set; }
        //{
        //    get
        //    {
        //        return CCAB_INS_DATE;
        //    }
        //    set
        //    {
        //        if (CAB_INS_DATE != null)
        //        {
        //            CCAB_INS_DATE = value;//(System.DateTime)CAB_INS_DATE;

        //        }
        //    }
        //}

        public string ALM_EXT { get; set; }
        public string ALM_EXT_TO { get; set; }
        public string ALM_STATUS { get; set; }


        DateTime CRC_WARRANTY_END = DateTime.Now;//set Default Value here
        [Display(Name = "RC_WARRANTY_END")]
        [DataType(DataType.Date)]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMMM-yyyy}")]
        public DateTime? RC_WARRANTY_END { get; set; }
        //{
        //    get
        //    {
        //        return CRC_WARRANTY_END = CRC_WARRANTY_END;
        //    }
        //    set
        //    {
        //        if (RC_WARRANTY_END != null)
        //        {
        //            CRC_WARRANTY_END = value;//(System.DateTime)RC_WARRANTY_END;

        //        }
        //    }
        //}

        public string BATT_EXIST { get; set; }
        DateTime CBATT_WARRANTY_END = DateTime.Now;//set Default Value here
        [Display(Name = "BATT_WARRANTY_END")]
        [DataType(DataType.Date)]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd-MMMM-yyyy}")]
        public DateTime? BATT_WARRANTY_END { get; set; }
        //{
        //    get
        //    {
        //        return CBATT_WARRANTY_END = CBATT_WARRANTY_END;
        //    }
        //    set
        //    {
        //        if (BATT_WARRANTY_END != null)
        //        {
        //            CBATT_WARRANTY_END = value;//(System.DateTime)BATT_WARRANTY_END;

        //        }
        //    }
        //}

        public string BATT_NO_CELL { get; set; }

        public string AIRCOND_IN_OUT { get; set; }
        public string AIRCOND_HPOWER { get; set; }
        public string REMARKS { get; set; }

        public string PTT { get; set; }
        public string EXCHANGE { get; set; }
        public string MODEL { get; set; }
        public string FACILITIES { get; set; }

        public string CODE { get; set; }
        public string DATA_NO { get; set; }

        public string[] leftValues { get; set; }
        public string[] rightValues { get; set; }

        public string STATUS { get; set; }
        public string SHARE_NO { get; set; }

        public string latlong { get; set; }
        public string INSTALL_YEAR { get; set; }
        public string BATT_EXTRA_LOCKING { get; set; }
    }
}
