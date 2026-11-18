"""
AI Survey Invite Helper
-----------------------
3 steps:  READ doctors from Excel  ->  WRITE a personal invite  ->  CHECK the invite
Result:   invites_output.xlsx  (each invite + PASS/FIX + what is wrong)

Run:      python3 invite_helper.py
AI mode:  set USE_AI = True  (needs: pip install anthropic  +  an Anthropic API key)
"""
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

HERE = Path(__file__).parent
INPUT_FILE = HERE / "doctors.xlsx"
OUTPUT_FILE = HERE / "invites_output.xlsx"
USE_AI = False  # False = simple template, True = Claude writes the emails


# ---------- STEP 1: READ ----------
def read_doctors(path):
    sheet = load_workbook(path).active
    headers = [cell.value for cell in sheet[1]]
    doctors = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        doctors.append(dict(zip(headers, row)))
    return doctors


# ---------- STEP 2: WRITE ----------
def write_with_template(d):
    return (
        f"Subject: Share your view on {d['Topic']} ({d['LOI']} min, ${d['Honorarium']})\n\n"
        f"Dear Dr. {d['Name']},\n\n"
        f"As a specialist in {d['Specialty']}, your opinion matters to us. "
        f"We invite you to a {d['LOI']}-minute survey on {d['Topic']}. "
        f"You will receive ${d['Honorarium']} as a thank-you.\n\n"
        f"Please complete it by {d['Deadline']}. Your answers are confidential.\n\n"
        "Thank you,\nCommunity Team"
    )


def write_with_ai(d):
    import anthropic  # only needed in AI mode

    client = anthropic.Anthropic()  # reads the API key from your environment
    prompt = (
        f"Write a short, polite survey invitation email (under 120 words) to Dr. {d['Name']}, "
        f"a {d['Specialty']} specialist. Topic: {d['Topic']}. Length: {d['LOI']} minutes. "
        f"Thank-you payment: ${d['Honorarium']}. Deadline: {d['Deadline']}. "
        "Start with a 'Subject:' line. Make no medical claims. Sign off as 'Community Team'."
    )
    response = client.beta.messages.create(
        model="claude-opus-5-5",
        max_tokens=1000,
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",  # if the model declines, the API retries on a fallback model
        messages=[{"role": "user", "content": prompt}],
    )
    if response.stop_reason == "refusal":
        return "AI declined to write this email."
    return "".join(block.text for block in response.content if block.type == "text")


# ---------- STEP 3: CHECK ----------
def check_email(email, d):
    problems = []
    if not d["Name"] or f"Dr. {d['Name']}" not in email:
        problems.append("Name wrong or missing")
    if not d["Honorarium"] or f"${d['Honorarium']}" not in email:
        problems.append("Payment wrong or missing")
    if not d["Deadline"] or str(d["Deadline"]) not in email:
        problems.append("Deadline wrong or missing")
    if len(email.split()) > 150:
        problems.append("Too long")
    return problems


# ---------- RUN ALL 3 STEPS ----------
def main():
    doctors = read_doctors(INPUT_FILE)

    out = Workbook()
    sheet = out.active
    sheet.title = "Invites"
    sheet.append(["Name", "Specialty", "Invite", "Check", "Problems", "Human approved? (Y/N)"])

    passed = 0
    for d in doctors:
        email = write_with_ai(d) if USE_AI else write_with_template(d)
        problems = check_email(email, d)
        status = "PASS" if not problems else "FIX"
        passed += status == "PASS"
        sheet.append([d["Name"], d["Specialty"], email, status, ", ".join(problems), ""])
        print(f"{status:4} | Dr. {d['Name']:8} | {', '.join(problems) or 'all checks OK'}")

    for col, width in zip("ABCDEF", [10, 16, 70, 8, 28, 20]):  # make the sheet readable
        sheet.column_dimensions[col].width = width
    for cell in sheet["C"]:
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    out.save(OUTPUT_FILE)
    print(f"\n{passed}/{len(doctors)} invites passed. Saved to {OUTPUT_FILE.name}")
    print("A human must review every invite before anything is sent.")


if __name__ == "__main__":
    main()
