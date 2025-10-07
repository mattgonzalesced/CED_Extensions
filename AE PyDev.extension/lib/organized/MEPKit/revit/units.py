# -*- coding: utf-8 -*-
FT_PER_M = 3.28083989501312

def ft(val):        return float(val)
def inches(val):    return float(val) / 12.0
def mm_to_ft(mm):   return float(mm) / 304.8
def ft_to_mm(ftv):  return float(ftv) * 304.8
def m_to_ft(m):     return float(m) * FT_PER_M
def ft_to_m(ftv):   return float(ftv) / FT_PER_M