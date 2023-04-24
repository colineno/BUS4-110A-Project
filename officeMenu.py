import pandas as pd

#Load DataFrame

x1 = pd.read_excel("TableauSalesData (1).xlsx")

#Start of option functions for the menu
#Option 1 function
def AverageProductDiscount(x1):
  #Calculating the average discount for all products
    return x1['Discount'].mean()

#Option 2 function
def TotalRegionDiscount(x1):
    regionDiscount = x1.groupby('Region')['Discount'].mean()
    #After getting the mean discounts for all regions, we load this into its own data frame
    regionDataFrame = pd.DataFrame({'Region': regionDiscount.index, 'Total Discounts': regionDiscount.values})
    
    return regionDataFrame

#Option 3 function
def RegionProfitDiscount(x1):
  #Finding the total profit and discount for each region
    RPD = x1.groupby('Region')[['Profit','Discount']].sum()
    RPD['Average Discount'] = RPD['Discount']/len((x1))

    TotalRegionProfit = RPD['Profit'].idxmax()
    #using the .loc technique
    averageDiscount = RPD.loc[TotalRegionProfit, 'Average Discount']
    return TotalRegionProfit, averageDiscount

#Option 4 Function
def BestSellingProduct(x1):
    productSold = x1.groupby('Product Name')['Sales', 'Profit'].sum()
    #here we used the idxmax panda function
    bestProduct = productSold['Sales'].idxmax()
    profit = productSold.loc[bestProduct, 'Profit']
    return f"The most relevant product among customers is {bestProduct} with a total profit of {profit:.2f}"

#Option 5 Function
def MostOrderedItem(x1):
  #calculating the discounts (sales - profit)
    x1['Discount'] = x1['Sales'] - x1['Profit']

#here we group the data
    dataGroup = x1.groupby('Product Name').agg({'Sales': 'sum', 'Profit': 'sum', 'Discount': 'sum', 'Quantity': 'sum'})
    MostOrderedProduct = dataGroup['Quantity'].idxmax()

    sales = dataGroup.loc[MostOrderedProduct, 'Sales']
    profit = dataGroup.loc[MostOrderedProduct, 'Profit']
    discount = dataGroup.loc[MostOrderedProduct, 'Discount']

    return MostOrderedProduct, sales, profit, discount

#Option 6 Function
def topTenCustomers(x1):
    
    customerSales = x1.groupby('Customer Name')['Sales'].sum()
    #sorting values to get the top 10 customers
    topCustomers = customerSales.sort_values(ascending=False)[:10]
    

    customerDiscounts = pd.merge(topCustomers, x1[['Customer Name', 'Discount']], on='Customer Name', how='left')
    
    #now we get the customer names with the avergae discount applied as well
    averageDiscount = customerDiscounts.groupby('Customer Name')['Discount'].mean()
    
    result = pd.concat([topCustomers, averageDiscount], axis=1)
    
    #put this data into columns
    result.columns = ['Total Sales', 'Average Discount']
    return result

#Start of the menu options
print("Welcome to the Office Solutions Data Analytics System")

def menu():
    print("\n Press 1 to see average discounts for all products"+
        "\n Press 2 to see the average regional discounts for each region"+
        "\n Press 3 to see most profited region with the average discounts applied"+
        "\n Press 4 to see the most relevant product among customers"+
        "\n Press 5 to see the most ordered item among customers "+
        "\n Press 6 to see the the top 10 customer who has the most orders with us "
        "\n Any other number options are invalid")
  
    userChoice = input("Enter what option you need: ")
    print("\n")
  
    if userChoice == "1":
        AverageProductDiscount(x1)
        x = AverageProductDiscount(x1)
        print("The average product discount for all products is:" , (x))
        menu()
    elif userChoice == "2":
        TotalRegionDiscount(x1)
        y = TotalRegionDiscount(x1)
        print("Average Regional Discount")
        print(y)
        menu()
    elif userChoice == "3":
        RegionProfitDiscount(x1)
        x = RegionProfitDiscount(x1)
        print('The most profitable Region is:', (x))
        menu()
    elif userChoice == "4":
        BestSellingProduct(x1)
        x = BestSellingProduct(x1)
        print(x)
        menu()
    elif userChoice == "5":
        MostOrderedItem(x1)
        x = MostOrderedItem(x1)
        print("Here is the total combined sales, profit, and discount for the most ordered item: ",(x))
        menu()
    elif userChoice == "6":
        topTenCustomers(x1)
        x = topTenCustomers(x1)
        print("Top Ten Customers")
        print(x)
    else:
        print("Invalid Input")
        menu()
menu()
