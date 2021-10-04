# VBA Stock Analysis

## Overview of Project
* Using VBA and Macros to provide Technical Analysis on Green Stocks 

### Purpose
* Taking our initial code for Technical Analysis on 12 similar stocks and taking yearly return and total volume for the past 2 years to identify and highlight any key metric and trends that provide insight on which stock is worth investing in

## Analysis and Challenges

### Outcomes based on Launch Date 
* This exploratory data analysis was performed on thousands of previous crowdfunding projects to highlight key metrics such as outcomes based on various factors such as funding goals, timelines, budgets, type of campaigns. 
* By taking the data set and using conditonal formatting and filters we can create pivot tables and pivot charts the highlight specific data that answers niche questions about these campaigns.  
* To do this I started with taking the data set and homogenizing it by freezing panes, applying filters and color coding the outcomes for successful, failed cancelled or live campaigns. After, creating new columns I created one for Years to see how each category has done over a yearly period. Taking the date conversion we created using Unix time stamp converters, we're able to focus on the years of these projects by using the Year( ) Function. 
* Once this has been done we take a Pivot Table using Years and Parent Category to see the outcome of all categories over a year
![Pivot Table for Outcomes by Launch Date](https://github.com/RichardH395/repo-kickstarter-analysis/blob/main/Pivot%20Table%20for%20Outcomes%20by%20Launch%20Date.png)
* We can visualize this to show the outcomes of all categories that show how theatre is a very popular category: 
![Parent Category Outcomes](https://github.com/RichardH395/repo-kickstarter-analysis/blob/main/Parent%20Category%20Outcomes%20.png) 
* As we can clearly see, Theatre is a popular category. Following up on this we filter categories to look at Theatre specifically to see a more detailed explanation regarding to Louise's question of how Launch dates affect these outcomes. We can take the pivot table and filter for theatre.
* ![Pivot Table for Theatre Outcomes by Launch Date](https://github.com/RichardH395/repo-kickstarter-analysis/blob/main/Pivot%20Table%20for%20Theatre%20Outcomes%20by%20Launch%20Date%20.png) 
#### Continuing Further
* Now that we have the data showing how outcomes are affected by launch date for Theatres, we create a pivot chart to better visualize this trend:
![Outcomes Based on Launch Date](https://github.com/RichardH395/repo-kickstarter-analysis/blob/main/Outcomes%20Based%20on%20Launch%20Date%20.png)
This is our first Deliverable, now we have a line chart with markers that show the outcome of the theatre category. We can see that May through July are popular launch dates for successful campaigns. 
### Outcomes based on Goals
* To gain insight on how much funding determines whether or not a campaign will be successful we narrow down our search to plays specifally. We create a column with different ranges of funding amounts to see what the range is for each outcome: successful, canceled, and failed and use the CountIf( ) to pull the Goal amounts within each range, see what the outcome was for each respective outcome, goal amount, and populate the spreadsheet. Following this we take the totals using the Sum( ) Function, after this we can take the percentage for every outcome and funding goal amount to see a summary of each outcome at each funding range. By taking the percentage we can create a Pivot table and chart showing the percentages on the Y-axis, and the funding amounts on the X-axis. We now have a template to showcase the percentage of each outcome whether failed or successful. Our findings are displayed here:
![Outcomes vs Goals](https://github.com/RichardH395/repo-kickstarter-analysis/blob/main/Outcomes%20vs%20Goals%20.png)


### Analysis of Outcomes Based on Launch Date
* As we can see from our findings and graphs, There are three Parent Categories that stand out in terms of both volume and success: Music, Technology, and Theatre. The Parent Category Outcomes Graph easily displays that Theatre is the most prominent Parent Cateogry both in terms of volume and outcomes. Having a large success rate, we can this take this analysis even further by highlighting theatre outcomes specifically. We find that the most successful months that theatres were launch are from May to July. Which is the peak in volume as well. We now have a window for when theatres are most succesful according to when they're launched.

### Analysis of Outcomes Based on Goals
* Now that we've concluded the peak months for successful theatre campaigns, we can investigate further by looking into the average amounts that these productions require according to their goals. We see that the most prominent range for goal amounts is between 1000-4999, very close to Louise's range for her funding amount. This entails almost half of all the plays; This is good news as it's a very feasible budget with a capacity of higher success. Espescially if they launch in the recommended time frame. We see as the amount increases, especially passed the 20,000 range,  we see a fall off in the number of campaigns but also the success rate drops drastically. This can be due to difficulty in recieving funding and support for such large plays and production. Less is more in this case and Louise can rest easy knowing she has all the checkmarks for a succesful kickstarter campaign for her play *Fever* 
### Challenges and Difficulties Encountered

After taking the data and formatting it, the difficulty can be in how you can use this data to answer specific questions. Setting up your framework for what you're looking for then to be able to take the data and present it in a way that tells a story can be very visually appealing and easy to consume but also hard to find, if you don't know what you're looking for. 

## Results

#### What are two conclusions you can draw about the Outcomes based on Launch Date?
* The Two conclusions are that Theatre is the most prominent Parent Category, we take this even further and find that the most successful theatre campaigns launch between May and July. 
#### What can you conclude about the Outcomes based on Goals?
* The conclusion from Outcomes based on Goals show that you don't need a large amount to fund a successful campaign, almost half of the campaigns range from 1000 to 4999. This a very financially feasible amount that has the opportunity to be very successful in the sense that less is more. As the amounts get larger they tend to become less successful and less likely to be funded. 

#### What are some limitations of this dataset?
* One of the limitation of the data set is that we're dealing with a sample amount of Kickstarter Campaigns, if we were to take every single campaign since Kickstarer's inception, it would provide a more accurate representation of of outcomes and goals for every category and call campaigns. Conveniently, this date set really focuses on theatres, which relates to Louise's inquiry in the first place. Looking historically it's important to see time periods where these categories and sub categories where thriving. This could be due to factors that are can or can not be accounted for when looking at these campaigns. For example: The Pandemic put a halt on Theatres and Plays, how does this get accounted for? Times of economic booms or bust, consumer trends and preference changing with techonoligcal and filming technique enhancements, or times when theatres were more prominent historically? 

* Another Limitation is the data set itself, it provides key metrics but some of these campaigns might or might not be more successful based on the premise of the outcome. What kind of rewards, recognition, concessions are offered for these campaigns? These kind of concessions can determine whether or not an individual decides to fund a campaign. On top of this looking at vast difference in plays. Low budget plays with small scripts and run times might be easier to fund and more likely to succeed, but what about large scale productions? There isn't much information on what funding these campaigns entail, more so what the objective is. 

#### What are some other possible tables and/or graphs that we could create?
* Tables and graphs that show ALL metrics and information for every single campaign would tell you about Kickstarter campaigns in general. To continue narrowing dowm, include every single campaign for every category, subcategory, funding, and outcome by compartmentalizing all of these you'd get very specific information regarding what makes each of these campaigns more or less successful and different levels of Funding Goals. Tables and Graphs showing the average time frame for every single category and sub category to get a benchmark of what each specific timeline every campaign should be following if they want to succeed. 
