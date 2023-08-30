#!/usr/bin/env python

from datetime import datetime
from io import BytesIO

from pandas import DataFrame, Series, read_csv
from reportlab.lib.colors import CMYKColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table

styles = getSampleStyleSheet()
styles["Heading1"].spaceBefore = 12
styles["Heading1"].spaceAfter = 0
styles["Heading1"].fontSize = 14

black = CMYKColor(0, 0, 0, 1)

SLIDO_EVENT_CODE = "rsecon23"
"""The code to access the main Q&A portal for the conference from slido.com"""
SLIDO_EVENT_HASH = "mwHnJYH3zt11GVLKnUTfvu"
"""The specific hash sequence that will get you directly to the event page or wall"""
SLIDO_EVENT_LINK = f"https://app.sli.do/event/{SLIDO_EVENT_HASH}"
"""The Q&A portal for the whole event. Specific room selection will be required."""
SLIDO_EVENT_WALL = f"https://wall.sli.do/event/{SLIDO_EVENT_HASH}"
"""The wall presentation page for the whole event. Specific room selection will be required."""
SLIDO_ROOM_WALL_MAP = {
    "GH043": "https://wall.sli.do/event/mwHnJYH3zt11GVLKnUTfvu?section=7c7a20b8-7014-4030-8e3a-08ae40150fa9",
    "GH049": "https://wall.sli.do/event/mwHnJYH3zt11GVLKnUTfvu?section=b7b4e005-652e-408e-a7dd-4458980bcfb0",
    "GH037": "https://wall.sli.do/event/mwHnJYH3zt11GVLKnUTfvu?section=a97ae336-63cb-461c-977c-ec5a03a06b12",
    "GH001": "https://wall.sli.do/event/mwHnJYH3zt11GVLKnUTfvu?section=7f16f6af-3588-4c53-980d-93880e132579",
}
"""The wall presentation page for each specific room."""


BASE_DOC = [
    "Dear {chair},",
    "Many thanks for agreeing to chair the <b>{session}</b> session at RSECon23. Please find below some guidance on your role as session chair. Please note that this guidance refers primarily to talks, but relevant adaptations are provided at the end for walkthroughs and panels.",
    ("Before the session", styles["Heading1"]),
    "Your session starts at <b>{start_time} on {day}</b>. Please arrive at room <b>{room}</b> 10 minutes in advance of your session. Your room will have two volunteers present and a stream director may be allocated if necessary, to help the event run smoothly—they will manage such things as moderating the Slido questions, ensuring speakers have microphones, loading up presentations onto the PC to display, and ensuring remote presenters are connected to the Zoom call and audible. Please introduce yourself to your volunteers before the start of the session.",
    "The speakers in your session should also be present before the start of the session to introduce themselves. <b>Please confirm the correct pronunciation of their names and pronouns</b>, especially if the name is unfamiliar to you.",
    "One of the room volunteers is responsible for logging onto the presenting PC. Ensure that the volunteer or stream director has the Slido wall view on one of the projected screens. The direct link <a href={slido_room_wall_link}><u>is here</u></a>. Alternatively, navigate to <a href={slido_event_wall}><u>{slido_event_wall}</u></a> and select the room {room}. It might drop you into the wrong room by default. Click next to the coloured dot at the top with text starting 'GH' to get a drop down of rooms.",
    ("Before each talk", styles["Heading1"]),
    "Please briefly introduce the speaker—a full bio is not needed, just the presenter's name and the title of the talk. Only do this after the buffer time between talks has elapsed and you have confirmed that the speaker is ready to start.",
    "Remind the audience that they can ask questions via Slido and point to the Slido wall slide. The manual procedure if not using the QR code is to navigate to <a href=https://slido.com>slido.com</a> and enter the event code: {slido_event_code}. Audience members will then need to select the correct room from the drop down. Remind them that it is {room}.",
    "Once the speaker begins, set a timer.",
    ("During each talk", styles["Heading1"]),
    "Listen to the talk and identify possible questions to ask.",
    "You can keep track of questions by going to the same Slido room as audience members just described above. The direct link to the event is <a href={slido_event_link}><u>{slido_event_link}</u></a> and you will still need to select the correct room.",
    "Talks are 20 minutes long, with 5 further minutes for questions. After 15 minutes, please show the “5 minutes remaining” sign to the speaker. Similarly, at 18, 19, and 20 minutes show the 2, 1, and 0 minutes remaining signs.",
    "If at 20 minutes the speaker shows no sign of concluding, then please stand up and look conspicuous. If after 30 seconds the speaker still shows no sign of concluding, then please interrupt them and ask them to briefly wrap up. (We have the luxury of a few minutes' buffer between talks. This is to enable people to move between rooms, and to set up the next speaker; it is not to enable speakers to overrun.)",
    ("After each talk", styles["Heading1"]),
    "After each talk is 5 minutes of Q&amp;A. The volunteer will switch the display to show the Slido wall view. All questions should be asked via Slido, do not take questions directly from the room.",
    "Read out the questions from Slido in descending order of upvotes. <b> Please be mindful not to stand between the speaker and the camera when reading out questions.</b> If there are no questions on Slido at the start of the Q&amp;A then please ask one or two questions yourself.",
    "Please do not take questions from the room unless and until questions from Slido have been exhausted.",
    "If participants in the room try to get into protracted discussions with the speaker during the Q&amp;A, please encourage them to use the coffee breaks to continue the discussion.",
    "Please keep an eye on the time during the Q&amp;A and close off questions in sufficient time to allow movement between rooms and setting up for the next speaker. Remind the volunteer to clear all Slido questions so they don't get mixed with those for the next presentation.",
    "At this point you can pre-announce the next speaker and the start time of their talk.",
]

PANEL_SECTION = [
    ("Panels", styles["Heading1"]),
    "Your session contains a panel session. Panels have their own chairs who take on some of the responsibilities above. Panels last 50 minutes, with 10 minutes slack to allow the panel to assemble and get ready at the start, and also to prepare for the next event at the end. However, due to scheduling constraints this buffer might not always be possible. Therefore, please ensure any talks before a panel end promptly to allow the panel to get ready.",
    "Once the panelists have assembled and have microphones ready, then please introduce the panel chair and topic similarly as for talks. Once the panel chair takes over and starts to introduce the panelists, please start a timer.",
    "After 45 minutes, please show the “5 minutes remaining” sign to the panel chair. Similarly, at 48, 49, and 50 minutes show the 2, 1, and 0 minutes remaining signs.",
    "If at 50 minutes the panel shows no signs of wrapping up, please stand up and look conspicuous. If after another minute the discussion or monologue is still continuing and the chair hasn't taken action, then please interrupt them and ask the chair to wrap things up.",
]

WALKTHROUGH_SECTION = [
    ("Walkthroughs", styles["Heading1"]),
    "You are chairing a Walkthrough session. Timings are slightly different than those for talks given above, although most of the other guidance holds true. Walkthroughs are 30 minutes, with 15 further minutes for questions. Please use the time remaining signs as described above but starting at 25 minutes. Please also note that walkthroughs are scheduled back-to-back with no buffer time.",
]

REMOTE_SECTION = [
    ("Remote presenters", styles["Heading1"]),
    "Your session includes a remote presenter. Your volunteer will launch Zoom on the PC in full screen mode, and the speaker will present via a shared screen. Please give the speaker audible notification of the remaining time, as they may not be looking at the camera feed from the room while presenting.",
]


def format_elements(template, session):
    elements = []
    for element in template:
        if isinstance(element, str):
            text = element
            style = styles["Normal"]
        else:
            text, style = element
        text = text.format(
            chair=session["Confirmed chair"]
            if isinstance(session["Confirmed chair"], str)
            else "chair",
            session=session["Session"],
            start_time=session["Session start time"],
            day=session["Day"],
            room=session["Room"],
            login_username=session["PC login username"],
            login_password=session["PC login password"],
            slido_event_code=SLIDO_EVENT_CODE,
            slido_event_link=SLIDO_EVENT_LINK,
            slido_event_wall=SLIDO_EVENT_WALL,
            slido_room_wall_link=SLIDO_ROOM_WALL_MAP[session["Room"]],
        )
        elements.append(Paragraph(text, style))
    return elements


def generate_infosheet(session: Series, talks: DataFrame, filename: str):
    buf = BytesIO()

    output_doc = SimpleDocTemplate(
        buf,
        rightMargin=30 * mm,
        leftMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
        pagesize=A4,
    )

    doc_contents = generate_infosheet_contents(session, talks)
    output_doc.build(doc_contents)

    with open(filename, "wb") as f:
        f.write(buf.getvalue())


def format_time(oa_time_string: str) -> str:
    """Format the time exported by Oxford Abstracts

    There was a bug in OA that meant the time string produced by an export was one hour
    later in the day. However, this bug has now been fixed, so this function is no
    longer used for that purpose, but simply to format the time correctly.

    Parameters
    ----------
    oa_time_string : str
        The time string produced by Oxford Abstracts

    Returns
    -------
    str
        The time string in HH:MM format
    """
    oa_time = datetime.strptime(oa_time_string.split(",")[2][6:], "%H:%M")
    return oa_time.time().strftime("%H:%M")


def tabulate(talks: DataFrame) -> Table:
    table_content = [["Start time", "End time", "Speaker", "Title", "Event type"]]
    for _, talk in talks.sort_values("Program individual start time").iterrows():
        if talk["Remote presentation"]:
            if talk["Event type"].startswith("Panel"):
                # Ugly hack to avoid text crashing but avoid consuming too much width
                remote_tag = "                    Remote panelist"
            else:
                remote_tag = "Remote speaker"
        else:
            remote_tag = ""
        table_content.append(
            [
                format_time(talk["Program individual start time"]),
                format_time(talk["Program individual end time"]),
                Paragraph(talk["Presenting"]),
                Paragraph(talk["Title"]),
                talk["Event type"],
                remote_tag,
            ]
        )

    table_style = [
        ("LINEABOVE", (0, 1), (-1, 1), 1, black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]

    table = Table(
        table_content, colWidths=[18 * mm, 18 * mm, 28 * mm, 65 * mm, 15 * mm, 15 * mm]
    )
    table.setStyle(table_style)
    table.hAlign = "CENTER"

    return table


def generate_infosheet_contents(session: Series, talks: DataFrame):
    doc_contents = format_elements(BASE_DOC, session)
    if any(talks["Event type"].str.startswith("Panel")):
        doc_contents.extend(format_elements(PANEL_SECTION, session))

    if session["Session"].count("Walkthrough"):
        doc_contents.extend(format_elements(WALKTHROUGH_SECTION, session))

    if talks["Remote presentation"].any():
        doc_contents.extend(format_elements(REMOTE_SECTION, session))

    doc_contents.append(Paragraph("Running order", styles["Heading1"]))
    doc_contents.append(tabulate(talks))

    return doc_contents


def get_filename(session):
    if isinstance(session["Confirmed chair"], str):
        filename = session["Confirmed chair"] + " - " + session["Session"] + ".pdf"
    else:
        filename = session["Session"] + ".pdf"

    return filename.replace("/", "_")


def get_talks(session: Series, talks: DataFrame) -> DataFrame:
    return talks[talks["Program session"] == session["Session"]]


def main():
    from argparse import ArgumentParser

    parser = ArgumentParser()
    parser.add_argument("sessions")
    parser.add_argument("talks")
    parser.add_argument("filename_prefix")
    args = parser.parse_args()

    sessions: DataFrame = read_csv(args.sessions, dtype={"Slido event code": str})
    talks: DataFrame = read_csv(args.talks)
    for _, session in sessions.iterrows():
        if not isinstance(session["Session"], str):
            continue
        generate_infosheet(
            session,
            get_talks(session, talks),
            args.filename_prefix + "/" + get_filename(session),
        )


if __name__ == "__main__":
    main()
