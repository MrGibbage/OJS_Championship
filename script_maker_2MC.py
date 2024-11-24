# pip install pandas
import pandas as pd

# pip install pyyaml
import yaml
import warnings

# pip install Jinja2
from jinja2 import Environment, FileSystemLoader, select_autoescape

SPEAKER_1: str="<p style=\"background-color:magenta;\"><b>Barry: </b>"
SPEAKER_2: str="<p style=\"background-color:aqua;\"><b>Andrew: </b>"
TEMPLATE_FILE: str = "script_template-2MC.html.jinja"
YAML_DATA_FILE: str = "meta.yaml"
AWARDS: dict[str, str] = {
    "Champions": "",
    "Innovation Project": "",
    "Robot Design": "",
    "Core Values": "",
}
ORDINALS: list[str] = ["1st", "2nd", "3rd", "4th", "5th"]

current_speaker: str = SPEAKER_1

# To create windows exe executable, run
# .venv\Scripts\pyinstaller.exe -F script_maker.py
# in the project folder. The executable will be saved in the 'dist'
# folder. Just copy it up to the project folder.
# Double-click to run.

print("Building Closing Ceremony script")

template_loader = FileSystemLoader(searchpath="./")
template_env = Environment(loader=template_loader)
try:
    template = template_env.get_template(TEMPLATE_FILE)
except:
    print(f"Template not found: {TEMPLATE_FILE}")
    quit()

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

print("Opening yaml data file: " + YAML_DATA_FILE)
try:
    with open(YAML_DATA_FILE) as f:
        dict = yaml.load(f, Loader=yaml.FullLoader)
except:
    print(f"There was an error opening the meta file: {YAML_DATA_FILE}")
    quit()

rg_html = {}
dataframes = {}
# divAwards key will be 1 and/or 2, and the value will be the award_html dictionaries
# yes, it is a dictionary of dictionaries
# I'm not proud of it
div_awards = {}
div_awards[1] = {}
div_awards[2] = {} 
# awardHtml key will be the award name, and the values will be the html
award_html = {}
advancing_df = {}
advancing_html = {}
award_counts = {}

divisions = [1, 2]
removing = []
for div in divisions:
    if dict["div" + str(div) + "_ojs_file"] is not None:
        try:
            ojs_filename = dict["div" + str(div) + "_ojs_file"]
            print(f"Reading OJS file {ojs_filename}")
            dataframes[div] = pd.read_excel(
                ojs_filename,
                sheet_name="Results and Rankings",
                header=1,
                usecols=[
                    "Team Number",
                    "Team Name",
                    "Max Robot Game Score",
                    "Robot Game Rank",
                    "Award",
                    "Advance?",
                ],
            )
        except Exception as e:
            print(f"There was a problem reading the excel OJS file: {ojs_filename}")
            print(repr(e))
            quit()
    else:
        # make a list of the elements to remove
        # I can't just remove the element from divisions[] while iterating
        # on it because the loops will end early
        removing.append(div)
        rg_html[div] = ""
        award_html[div] = ""
        div_awards[div]["Robot Design"] = ""
        div_awards[div]["Innovation Project"] = ""
        div_awards[div]["Core Values"] = ""
        div_awards[div]["Champions"] = ""
        advancing_html[div] = ""

for award in AWARDS:
    award_counts[award] = 0

# now that we are done looping over the divisions, remove any items
# that were marked for deletion
# print(removing)
for item in removing:
    divisions.remove(item)
print("We have these divisions: " + str(divisions))
print("Get the top robot game scores")
# print(divisions)
# Robot Game
for div in divisions:
    rg_html[div] = ""
    print("Getting top scores for division " + str(div))
    for i in reversed(range(int(dict["Division " + str(div) + " Robot Game"]))):
        print(ORDINALS[i] + " Place")
        try:
            team_num = int(
                dataframes[div].loc[
                    dataframes[div]["Robot Game Rank"] == i + 1, "Team Number"
                ].iloc[0]
            )
            team_name = dataframes[div].loc[
                dataframes[div]["Robot Game Rank"] == i + 1, "Team Name"
            ].iloc[0]
            score = int(
                dataframes[div].loc[
                    dataframes[div]["Robot Game Rank"] == i + 1, "Max Robot Game Score"
                ].iloc[0]
            )
        except Exception as e:
            print(f"There was an error getting the robot game scores for Division {div}")
            print(repr(e))
            quit()

        current_speaker = SPEAKER_1 if current_speaker == SPEAKER_2 else SPEAKER_2
        rg_html[div] += (
            current_speaker
            + "With a score of "
            + str(score)
            + " points, the Division " + str(div) + " "
            + ORDINALS[i]
            + " place award goes to team number "
            + str(int(team_num))
            + ", "
            + team_name
            + "</p>\n"
        )

# print(rg_html)
# Judged awards
for div in divisions:
    div_awards[div] = {}
    print("Getting awards for division " + str(div))
    for award in AWARDS.keys():
        print("Getting the " + award + " award")
        div_awards[div][award] = ""
        award_counts[award] += int(dict["Division " + str(div) + " " + award])
        for i in reversed(range(int(dict["Division " + str(div) + " " + award]))):
            # divAwards[div][award] = ""
            print(award + " " + ORDINALS[i] + " Place")
            try:
                team_num = int(
                    dataframes[div].loc[
                        dataframes[div]["Award"] == award + " " + ORDINALS[i],
                        "Team Number",
                    ].iloc[0]
                )
                team_name = dataframes[div].loc[
                    dataframes[div]["Award"] == award + " " + ORDINALS[i],
                    "Team Name",
                ].iloc[0]
            except Exception as e:
                print(f"There was an error getting the {award} award for Division {div}")
                print(repr(e))
                quit()

            current_speaker = SPEAKER_1 if current_speaker == SPEAKER_2 else SPEAKER_2
            this_text = (
                current_speaker
                + "The Division " + str(div) + " "
                + ORDINALS[i]
                + " place " + award + " award goes to team number "
                + str(int(team_num))
                + ", "
                + team_name
                + "</p>\n"
            )
            div_awards[div][award] += this_text

print("Getting the advancing teams")
# Advancing
for div in divisions:
    try:
        advancing_df[div] = dataframes[div][dataframes[div]["Advance?"] == "Yes"][
            ["Team Number", "Team Name"]
        ]
    except Exception as e:
        print(f"There was an error getting the advancing teams for Division {div}")
        print(repr(e))
        quit()

# print(advancingDf[1])
# print(advancingDf[2])

# # From https://stackoverflow.com/questions/18695605/how-to-convert-a-dataframe-to-a-dictionary
# div1advancingDict = div1advancingDf.set_index("Team Number").to_dict("dict")
# div2advancingDict = div2advancingDf.set_index("Team Number").to_dict("dict")
# # print(div2advancingDict["Team Name"])

for div in divisions:
    advancing_html[div] = ""
    # print(advancingDf[div])
    for index, row in advancing_df[div].iterrows():
        try:
            team_num = str(int(row['Team Number']))
        except:
            team_num = ""
        team_name = row["Team Name"]
        current_speaker = SPEAKER_1 if current_speaker == SPEAKER_2 else SPEAKER_2
        advancing_html[div] += current_speaker + "(Div " + str(div) + ") Team number " + team_num + ", " + team_name + "</p>\n"

ja_count = int(dict["Judges Awards"])

print("Rendering the script")
out_text = template.render(
    tournament_name = dict["tournament_name"], 
    volunteer_award_justification = dict["Volunteer Justification"],
    volunteer_awardee_name = dict["Volunteer Awardee"],
    rg_div1_list = rg_html[1],
    rg_div2_list = rg_html[2],
    rd_div1_list = div_awards[1]["Robot Design"],
    rd_div2_list = div_awards[2]["Robot Design"],
    rd_this_them = "This team" if award_counts["Robot Design"] == 1 else "These teams",
    ip_div1_list = div_awards[1]["Innovation Project"],
    ip_div2_list = div_awards[2]["Innovation Project"],
    ip_this_them = "This team" if award_counts["Innovation Project"] == 1 else "These teams",
    cv_div1_list = div_awards[1]["Core Values"],
    cv_div2_list = div_awards[2]["Core Values"],
    cv_this_them = "This team" if award_counts["Core Values"] == 1 else "These teams",
    ja_count = ja_count,
    ja_list = str(dict["Judges Awardees"]),
    ja_go_goes = "The Judges Awards go to teams" if ja_count > 1 else "The Judges Award goes to team",
    special_guests = str(dict["Special Guests"]),
    champ_div1_list = div_awards[1]["Champions"],
    champ_div2_list = div_awards[2]["Champions"],
    adv_div1_list = advancing_html[1],
    adv_div2_list = advancing_html[2],
    spkr1 = SPEAKER_1,
    spkr2 = SPEAKER_2
)

# print(out_text)
with open(dict["complete_script_file"], "w") as fh:
    fh.write(out_text)

print("All done! The script hase been saved as " + dict["complete_script_file"] + ".")
print("It is saved in the same folder with the other OJS files.")
print("Double-click the script file to view it on this computer,")
print("or email it to yourself and view it on a phone or tablet.")
input("Press enter to quit...")