prices = [2,1,4,5,2,9,7]

def maxProfit(prices: list[int]) -> int:
    left = 0
    right = 1
    totalprofit = 0
    profit = 0
    while right<len(prices):
        if prices[right]<=prices[right-1]:
            left = right
            right = left + 1
            totalprofit += profit
            print('if statement current left/right: ',left,right,'total profit is ',totalprofit)
            profit = 0
            continue
        else:
            profit = prices[right]-prices[left]
            right+=1
            print('else statement current left/right: ',left,right,'temp profit is ',profit)
    return totalprofit + profit    
print(maxProfit(prices))
        
