import matplotlib.pyplot as plt
import numpy as np
from scipy import stats
from statsmodels.formula.api import ols
from statsmodels.stats.anova import anova_lm
import pandas
import xlrd


#Open Excel WorkBook --> (row,col)
wb = xlrd.open_workbook("charity.xlsx")
sheet = wb.sheet_by_index(0)


#Gift Variable
gift = []
for i in range(sheet.nrows):
    gift.append(sheet.cell_value(i,2))
gift.remove("GIFT")
gift = np.array(gift)


#Gift Descriptive Data
print("Gift mean:",np.mean(gift),
      "\nGift Std Dev:", np.std(gift),
      "\nGift Sample:", len(gift),
      "\nGift Max:", np.max(gift),
      "\nGift Min:", np.min(gift),
      "\n"
      )


#Gift histogram
plt.hist(gift, bins = 30, color = "skyblue", ec ="black")
plt.title("Gift")
plt.show()


#No gift pct Variable
no_gift = []
for i in range(len(gift)):
    if gift[i] == 0:
      no_gift.append(gift[i])
no_gift_pct = round(len(no_gift) * 100 / len(gift),2)
print("No gift pct:", str(no_gift_pct) + "%")


#GiftLast Variable
gift_last = []
for i in range(sheet.nrows):
    gift_last.append(sheet.cell_value(i,7))
gift_last.remove("GIFTLAST")
gift_last = np.array(gift_last)


#Propresp Variable
propresp = []
for i in range(sheet.nrows):
    propresp.append(sheet.cell_value(i,5))
propresp.remove("PROPRESP")
propresp = np.array(propresp)


#Mailsyear Variable
msy = []
for i in range(sheet.nrows):
    msy.append(sheet.cell_value(i,6))
msy.remove("MAILSYEAR")
msy = np.array(msy)


#Mailsyear Descriptive Data
print("Mailsyear mean:",np.mean(msy),
      "\nMailsyear Std Dev:", np.std(msy),
      "\nMailsyear Sample:", len(msy),
      "\nMailsyear Max:", np.max(msy),
      "\nMailsyear Min:", np.min(msy),
      "\n"
      )


#Mailsyear Histogram
plt.hist(msy, bins = 20, color = "skyblue", ec ="black")
plt.title("MAISYEAR")
plt.show()


#Linear Regression: gift = a + b*mailsyear
slope, intercept, r_value, p_value, std_err = stats.linregress(msy, gift)
print("Slope:", slope,
      "\nIntercept:", intercept,
      "\nR Squared:", pow(r_value,2),
      "\nStd Error:", std_err,
      "\n"
      )


#Linear Regression Plot
ec = intercept + slope * msy
plt.scatter(msy, gift, color = "blue")
plt.plot(msy, ec, color = "orange")
plt.xlabel("MAILSYEAR")
plt.ylabel("GIFT")
plt.title("Linear Model", fontsize =14)
plt.show()


#Multiple Regression
data = pandas.DataFrame({
    "mailsyear": msy,
    "giftlast": gift_last,
    "propresp": propresp,
    "gift": gift
    })
model = ols("gift ~ mailsyear + giftlast + propresp", data).fit()
print(model.summary())
