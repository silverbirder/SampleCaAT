# SampleCaAT
This repository is  trying the CaAT repotistory.

# Smaple Results

Bob's next week's schedule is:
![sampleCaAT](https://res.cloudinary.com/silverbirder/image/upload/v1579941127/CaAT/sampleCaAT.png)
Bob wants to know the scheduled **assigned time** next week.
However, Bob wants to exclude the time of **morning meeting, lunch time, and concentration time**.

When this scripts(src/index.ts main) is executed, the following result is obtained.
![sampleCaAT_spreatsheet](https://res.cloudinary.com/silverbirder/image/upload/v1579941128/CaAT/sampleCaAT_spreatsheet.png)

|Event|Assign Time(h)|Reason|
| --- | --- | --- |
|Morning MTG|0|Target the cut time|
|Lunch|0|Target the cut time|
|Marketing MTG|0|Event status is non-attendance|
|Tech MTG|0|Bob is morning break in 27 day|
|SEO MTG|0|Bob is all holiday in 31 day|
|Team Planning MTG|0|Bob is all holiday in 31 day|
|Concentration|0|Event is ignore|
|DepOps MTG|0.5|None|

※ 5 hours working per day. (2 hours is morning, 3 hours is afternoon)