#!/usr/bin/env python

from datetime import datetime, timedelta

from io import BytesIO

from pandas import read_csv

from reportlab.platypus import SimpleDocTemplate, PageBreak, Paragraph, Table
from reportlab.lib.colors import CMYKColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm

styles = getSampleStyleSheet()
styles["Heading1"].spaceBefore = 12
styles["Heading1"].spaceAfter = 0
styles["Heading1"].fontSize = 14

black = CMYKColor(0, 0, 0, 1)


BASE_DOC = [
    "Dear {chair},",
    "Many thanks for agreeing to chair the <b>{session}</b> session at RSECon2022. Please find below some guidance on your role as session chair.",
    ("Before the session", styles["Heading1"]),
    "Your session starts at <b>{start_time} on {day}</b>. Please arrive at room <b>{room}</b> 10 minutes in advance of your session. Your room will have a volunteer present to help the event run smoothly—they will manage such things as moderating the Sli.do questions, ensuring speakers have microphones, and loading up presentations onto the PC to display. Please introduce yourself to your volunteer before the start of the session.",
    "The speakers in your session should also be present before the start of the session to introduce themselves. <b>Please confirm the correct pronunciation of their names</b>, especially if the name is unfamiliar to you.",
    "In case you need it, the username for the PC is <b>{login_username}</b>, and the password is <b>{login_password}</b>.",
    ("Before each talk", styles["Heading1"]),
    "Please briefly introduce the speaker—a full bio is not needed, just the presenter’s name and the title of the talk. Remind the audience that they can ask questions via Sli.do.",
    "Once the speaker begins, set a timer.",
    ("During each talk", styles["Heading1"]),
    "Listen to the talk and identify possible questions to ask.",
    "You can keep track of questions by going to Sli.do room <b>{slido_room_number}</b>.",
    "Talks are 20 minutes long, with 5 further minutes for questions. After 15 minutes, please show the “5 minutes remaining” sign to the speaker. Similarly, at 18, 19, and 20 minutes show the 2, 1, and 0 minutes remaining signs.",
    "If at 20 minutes the speaker shows no sign of concluding, then please stand up and look conspicuous. If after 30 seconds the speaker still shows no sign of concluding, then please interrupt them and ask them to briefly wrap up. (We have the luxury of a few minutes’ buffer between talks. This is to enable people to move between rooms, and to set up the next speaker; it is not to enable speakers to overrun.)",
    ("After each talk", styles["Heading1"]),
    "After each talk is 5 minutes of Q&amp;A. The volunteer will switch the display to show Sli.do. Invite the audience to ask questions and upvote others’ questions at the link on screen.",
    "Read out the questions from Sli.do in descending order of upvotes. If there are no questions on Slido at the start of the Q&amp;A then please ask one or two questions yourself.",
    "Please do not take questions from the room unless and until questions from Slido have been exhausted.",
    "If participants in the room try to get into protracted discussions with the speaker during the Q&amp;A, please encourage them to use the coffee breaks to continue the discussion.",
    "Please keep an eye on the time during the Q&amp;A and close off questions in sufficient time to allow movement between rooms and setting up for the next speaker.",
    "At this point you can pre-announce the next speaker and the start time of their talk.",
]

PANEL_SECTION = [
    ("Panels", styles["Heading1"]),
    "Panels have their own chairs who take on some of the responsibilities above. Panels last 50 minutes, with 10 minutes slack to allow the panel to assemble and get ready at the start, and also to prepare for the next event at the end.",
    "Once the panelists have assembled and have microphones ready, then please introduce the panel chair and topic similarly as for talks. Once the panel chair takes over and starts to introduce the panelists, please start a timer.",
    "After 45 minutes, please show the “5 minutes remaining” sign to the panel chair. Similarly, at 48, 49, and 50 minutes show the 2, 1, and 0 minutes remaining signs.",
    "If at 50 minutes the panel shows no signs of wrapping up, please stand up and look conspicuous. If after another minute the discussion or monologue is still continuing and the chair hasn’t taken action, then please interrupt them and ask the chair to wrap things up.",
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
            slido_room_number=session["Slido room number"],
        )
        elements.append(Paragraph(text, style))
    return elements


def generate_infosheet(session, talks, filename):
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


def fix_time(oa_time_string):
    incorrect_time = datetime.strptime(oa_time_string.split()[3][:-3], "%H:%M")
    correct_time = incorrect_time - timedelta(hours=1)
    return correct_time.time().strftime("%H:%M")


def tabulate(talks):
    table_content = [["Start time", "End time", "Speaker", "Title", "Event type"]]
    for _, talk in talks.sort_values("Program submission start time").iterrows():
        table_content.append(
            [
                fix_time(talk["Program submission start time"]),
                fix_time(talk["Program submission end time"]),
                Paragraph(talk["Presenting"]),
                Paragraph(talk["Title"]),
                talk["Event type"],
                "Remote speaker" if talk["Remote presentation"] else "",
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


def generate_infosheet_contents(session, talks):
    doc_contents = format_elements(BASE_DOC, session)
    if session["Has panel"] == "Yes":
        doc_contents.extend(format_elements(PANEL_SECTION, session))
        doc_contents.append(PageBreak())

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


def get_talks(session, talks):
    return talks[talks["Program session"] == session["Session"]]


def main():
    from argparse import ArgumentParser

    parser = ArgumentParser()
    parser.add_argument("sessions")
    parser.add_argument("talks")
    parser.add_argument("filename_prefix")
    args = parser.parse_args()

    sessions = read_csv(args.sessions, dtype={"Slido room number": str})
    talks = read_csv(args.talks)
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
