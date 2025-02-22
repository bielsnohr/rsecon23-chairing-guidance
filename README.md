# Chairing guidance generator for RSECon23

This tool generates personalised guidance for session chairs for RSECon23.

## Setup

The file `environment.yml` contains a Conda environment specification; a corresponding environment can be created using:

    conda env create -f environment.yml

Two files are needed:

* An export of talks from Oxford Abstracts (assumed by the `Makefile` to be called `talks.csv`), with minimally the following columns enabled:
  * `Program session`
  * `Title`
  * `Presenting`
  * `Event type`
  * `Program submission start time`
  * `Program submission end time`
  * `Remote presentation` (a hidden checkbox field to allow the committee to track remote speakers)
* A CSV containing data about the sessions (assumed by the `Makefile` to be called `sessions.csv`), containing minimally the following columns:
  * `Session` - the name of the session
  * `Confirmed chair` - the name of the chair
  * `Session start time`
  * `Day` - the day on which the session occurs
  * `Room` - the room in which the session is scheduled
  * `PC login username` - the username to log in to the PC used to present
  * `PC login password` - the password to log in to the PC used to present

The information about Slido is specified within `chairing.py` in global variables prefixed with
`SLIDO_*`. Modify those as needed, or add another column in the `sessions.csv` and handle that in
the relevant sections of `chairing.py`.

## Usage

To activate the conda environment:

    conda activate chairing

To run the tool:

    make

This will place PDFs of guidance for each session into a directory called `infosheets`.

## Past Quirks

In a previous iteration, Oxford Abstracts got confused over time zones and reported the submission
start and end times offset by one hour during British Summer Time. This bug has been fixed, but it
is worth being aware of in case it pops up again. Modify the `format_time` function and look at `git
blame` to see how it was handled in the past.
