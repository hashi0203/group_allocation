# Group allocation
This repository is for group allocation with a lot of restrictions.<br>
By this program, you can make group of 4 three times, group of 3 twice, and group of 2 once.

# Dependency
Python 3<br>
Implement the necessary libralies with the command below.
```bash
$ pip install -r requirements.txt
```

# Usage
Clone this repository and execute make_groups.py.
```bash
$ git clone git@github.com:hashi0203/group_allocation.git
$ cd group_allocation
$ python3 make_groups.py
```

# Restrictions
1. Every group has at most 2 members in the same year.
2. Every member is not in the same group with others more than once during whole 8 times.
3. Every member is marked grey in output.xlsx once in each group of 4 and 3.
4. In group of 2, every member is in pair with another member in a different year.
5. In group of 2, the year difference in every pair is lower than 3.

# Input and Output Files
In every csv file, there are school years in collumn 0 and corresponding names in collumn 1.<br>
In memlist.csv, there are all members' years and names.<br>
In absentee.csv, there are all absent members' years and names.<br>
In early_leave.csv, there are all early-leavers' years and names.<br>
In output.xlsx, groups which meet the restrictions above are shown.

# Note
This is only for my case, so it may be difficult for you to use this for your cases.

