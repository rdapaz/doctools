=(SUM(Forecast!G23:AD23+SUM(Forecast!G44:AD44)-G11-G12))*HRP_Implementation_ProRata
=(SUM(Forecast!G109:AD109)+SUM(Forecast!G123:AD123)-G14-G15*HRP_Hardware_ProRata

=(SUM(Forecast!G23:AD23+SUM(Forecast!G44:AD44)-G11-G12))*(1-HRP_Implementation_ProRata)
=(SUM(Forecast!G109:AD109)+SUM(Forecast!G123:AD123)-G14-G15*(1-HRP_Hardware_ProRata)





=(H11+H12)*(1-HRP_Implementation_ProRata)
=(H13+H14)*(1-HRP_Hardware_ProRata)