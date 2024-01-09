For this assignment I followed the same logic as on the CreditCardChecker exercise for students on gitlab, I also reference to the exercise UScensus part1 and part 2. To loop through the different worksheets I went throught the documentation available on the US census part 1 student exercise where it explains how to loop into different worksheets at once and I used the same fundamentals found here "https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0". I also attended a tutoring session in which the tutor helped me to clean out my code seems it did not look very clean. I did ask XpertLearningAssistant for some guidance to find the yearly change and calculate the percentage change as accomodated it on my project.

    "  ' Calculate the yearly change
        Dim yearlyChange As Double
        yearlyChange = closingPrice - openingPrice
        
        ' Calculate the percentage change
        Dim percentChange As Double
        If openingPrice <> 0 Then
            percentChange = (yearlyChange / openingPrice) * 100
        Else
            percentChange = 0
        End If "
        
