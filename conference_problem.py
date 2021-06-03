"""
The Conference Themes-Speakers decision problem

Several themes could be presented during a conference
A set of potential speakers express their interest for these themes to be preseted
Which themes should be selected?
Who should present them?

We use PuLP Modeller to model and solve the problem

Author: Alessandro Giassi 2021
"""

# Packages
from pulp import *
import pandas as pd


# Main paremeters

max_themes = 6	# maximum number of themes to be selected among those listed

max_speakers_per_theme = 2	# maximum number of speakers presenting a selected theme

max_themes_per_speaker = 1	# maximum number of themes presented by a speaker

datafile = 'ThemesConferenceIA_31mai.xlsx'	# file where all the preferences are listed in a table

resultfile = 'Results.xlsx'	# file where results will be stored


#Declare the problem name and type
myproblem = LpProblem("The ConferenceIA Problem", LpMaximize)


#Read preferences from the form Excel file
form_data = pd.read_excel(datafile,skiprows=1,index_col=0)


#Extract the Themes from the rows names (index)
Themes = form_data.index.tolist()


#Extract the speakers names from the colums
Speakers = form_data.columns.tolist()


#Extract the interest of each speaker for a given theme ant put it into a list
interests = form_data.fillna(0).values.tolist()

#Pulp uses a function makeDict to transform this list into a dictionary (but it sort all indexes alphabeticaly)
interests = makeDict([Themes,Speakers],interests,0)

#Better sort also our lists Themes and Speakers
Themes.sort()
Speakers.sort()

#Create all possible attributions of a theme to a speaker
Attributions = [(t,s) for t in Themes for s in Speakers]

#Declare our main decision variables : 1 if the theme t is attributed to speaker s, 0 if not
is_theme_speaker = LpVariable.dicts("Attribution",(Themes,Speakers),0,1,LpInteger)

#Declare a secondary decision variable : 1 if a theme will be presented by someone, 0 if not
is_theme = LpVariable.dicts("Theme",Themes,0,1,LpInteger)

# The objective function is the sum of the binary decision variables multiplied by the interest scores  
# The objective function is added to 'myproblem' first
myproblem += lpSum([is_theme_speaker[t][s]*interests[t][s] for (t,s) in Attributions]), "Sum_of_Interests"

# Then we add constraints, here we want select only a number of themes <= max_themes
myproblem += lpSum([is_theme[t] for t in Themes])<=max_themes, "Sum_of_themes_selection"

# A constraints that each speaker is presenting no more than 'max_themes_per_speaker' themes
for s in Speakers:
    myproblem += lpSum([is_theme_speaker[t][s] for t in Themes])<=max_themes_per_speaker, "Sum_of_attribution_of_speaker_%s"%s

# A constraints that each theme is presented by no more than 'max_speakers_per_theme' speaker
for t in Themes:
    myproblem += lpSum([is_theme_speaker[t][s] for s in Speakers])<=max_speakers_per_theme*is_theme[t], "Sum_of_Speakers_for_theme_%s"%t

# The problem data is written to an .lp file
myproblem.writeLP("ConferenceIA.lp")

# The problem is solved using PuLP's choice of Solver
myproblem.solve()

# The status of the solution is printed to the screen
print("Status:", LpStatus[myproblem.status])

# Each of the variables is printed with it's resolved optimum value
#for v in myproblem.variables():
#    print(v.name, "=", v.varValue)

# The optimised objective function value is printed to the screen    
print("Total interest = ", value(myproblem.objective))


#Write the results into a data frame, then into an Excel file
results = pd.DataFrame(columns=Speakers,index=Themes)

for a in Attributions:
    results.loc[a] = myproblem.variables()[Attributions.index(a)].varValue
    
results.to_excel(resultfile)