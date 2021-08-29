
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import datetime
from bikram import samwat


class GenesisDate:
    """
    utilities for Genesis overall Dateobjects
    input : pandas.Series-Date columns
    output:
    -get_ad_date: 
    -get_ad_year
    -to_bs
    -get_month
    -get_quarter
    -get_day
    -get_fy
    -check_ad_with_bs: not yet updated
    """
    calendar={
        
    1:('q1','baishak'),2:('q1','jestha'),3:('q1','asar'),
    4:('q2','shrawan'),5:('q2','bhadra'),6:('q2','asoj'),
    7:('q3','kartik'),8:('q3','mangshir'),9:('q3','poush'),
    10:('q4','magh'),11:('q4','falgun'),12:('q4','chaitra')
        }

    
    def __init__(self,ad:list):
        ## need to run validation before this to ensure we are parsing dates
        ## pandas has more flexibility with datetime object there for converting every thing in pandas series
        ad=pd.Series(ad)
        if ad.astype(str).str.isdigit().sum()>0:
            f=ad.astype(str).str.isdigit()
            ad[f]=ad[f].astype(int).apply(self.from_excel_ordinal)
    
        self.ad=pd.to_datetime(ad) ## careful this is a date and time object; not the date object
        self.bs=pd.Series([samwat.from_ad(date) for date in self.ad.dt.date]) ## its better to store bs date in samwat object
        
    def __repr__(self):
        return ('genesis_date_objects')
   
    def from_excel_ordinal(self,ordinal, _epoch0=datetime.datetime(1899, 12, 31)):
        # https://stackoverflow.com/questions/29387137/how-to-convert-a-given-ordinal-number-from-excel-to-a-date
        ## for date formats like 4663,46384
        if ordinal >= 60:
            ordinal -= 1  # Excel leap year bug, 1900 is not a leap year!
        return (_epoch0 + timedelta(days=ordinal)).replace(microsecond=0)
        
    def get_ad_date(self):
        return self.ad.dt.date
    
    def get_ad_year(self):
        return self.ad.dt.year
    
    def to_bs(self):
        bs_date=[date.replace('-','.') for date in self.bs.astype(str)]
        return pd.Series(bs_date)
    
    def get_month(self):
        month=[self.calendar[date.month][1] for date in self.bs]  ## samwat object have .month for month
        return pd.Series(month)
    
    def get_quarter(self):
        quarter=[self.calendar[date.month][0] for date in self.bs]
        return pd.Series(quarter)
    
    def get_day(self):
        day=pd.to_datetime(self.ad).dt.day_name()
        return pd.Series(day)
    
    def get_fy(self):
        bins_dt = pd.date_range('2011-07-16', freq='365D', periods=20)
        bins_samwat=[str(samwat.from_ad(date).year) for date in bins_dt.date]

        bins_str = bins_dt.astype(str).values
        labels = [f'{bins_samwat[i-1]}-{int(bins_samwat[i])%100}' for i in range(1, len(bins_samwat))] ## modulo to get last two digit
        ## assin Fiscal years
        ## supports only datetime objects; thus excute this before converting to date
        return pd.cut(self.ad.view(np.int64)//10**9, ## unix epoch coversion
                       bins=bins_dt.view(np.int64)//10**9,
                       labels=labels)

    
