from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas


def draw_label(c, text, x, y, size=10):
    c.setFont("Helvetica-Bold", size)
    c.drawString(x, y, text)


def draw_paragraph(c, lines, x, y_start, line_height=14, size=10):
    c.setFont("Helvetica", size)
    y = y_start
    for line in lines:
        c.drawString(x, y, line)
        y -= line_height
    return y


def text_field(c, name, x, y, w, h=18, value="", multiline=False):
    c.acroForm.textfield(
        name=name,
        x=x,
        y=y,
        width=w,
        height=h,
        value=value,
        borderStyle="inset",
        forceBorder=True,
        fontName="Helvetica",
        fontSize=10,
        fieldFlags="multiline" if multiline else "",
    )


def make_pdf(output_path):
    c = canvas.Canvas(output_path, pagesize=LETTER)
    width, height = LETTER
    left = 40
    right = width - 40

    # Header
    c.setFont("Helvetica-Bold", 18)
    c.drawString(left, height - 45, "Freelance Project Proposal")
    c.setFont("Helvetica", 10)
    c.drawRightString(right, height - 42, "Fillable Template")

    y = height - 80
    draw_label(c, "Client:", left, y)
    text_field(c, "client_name", left + 60, y - 4, 220)
    draw_label(c, "Project:", left + 300, y)
    text_field(c, "project_name", left + 355, y - 4, 170)

    y -= 30
    draw_label(c, "Date:", left, y)
    text_field(c, "proposal_date", left + 60, y - 4, 140)
    draw_label(c, "Prepared by:", left + 220, y)
    text_field(c, "prepared_by", left + 300, y - 4, 225)

    y -= 34
    draw_label(c, "Project Overview", left, y)
    y -= 16
    text_field(c, "project_overview", left, y - 82, right - left, 86, multiline=True)

    y -= 100
    draw_label(c, "Scope of Work (bullet-friendly)", left, y)
    y -= 16
    text_field(c, "scope_of_work", left, y - 102, right - left, 106, multiline=True)

    y -= 122
    draw_label(c, "Deliverables", left, y)
    y -= 16
    text_field(c, "deliverables", left, y - 68, right - left, 72, multiline=True)

    y -= 90
    draw_label(c, "Out of Scope", left, y)
    y -= 16
    text_field(c, "out_of_scope", left, y - 60, right - left, 64, multiline=True)

    # Footer page 1
    c.setFont("Helvetica", 9)
    c.drawRightString(right, 18, "Page 1 of 3")
    c.showPage()

    # Page 2
    y = height - 45
    c.setFont("Helvetica-Bold", 14)
    c.drawString(left, y, "Timeline, Investment, and Terms")
    y -= 30

    draw_label(c, "Estimated Start Date:", left, y)
    text_field(c, "start_date", left + 135, y - 4, 120)
    draw_label(c, "Estimated End Date:", left + 275, y)
    text_field(c, "end_date", left + 400, y - 4, 125)

    y -= 28
    draw_label(c, "Milestones", left, y)
    y -= 16
    text_field(c, "milestones", left, y - 72, right - left, 76, multiline=True)

    y -= 94
    draw_label(c, "Pricing Option A (Fixed Fee)", left, y)
    y -= 16
    draw_label(c, "Total Project Fee:", left, y)
    text_field(c, "fixed_fee_total", left + 120, y - 4, 160)

    y -= 28
    draw_label(c, "Pricing Option B (Hourly)", left, y)
    y -= 16
    draw_label(c, "Hourly Rate:", left, y)
    text_field(c, "hourly_rate", left + 80, y - 4, 100)
    draw_label(c, "Estimated Hours:", left + 205, y)
    text_field(c, "estimated_hours", left + 310, y - 4, 100)
    draw_label(c, "Estimated Total:", left + 425, y)
    text_field(c, "hourly_estimated_total", left + 520, y - 4, 50)

    y -= 36
    draw_label(c, "Payment Terms", left, y)
    y -= 16
    text_field(c, "payment_terms", left, y - 72, right - left, 76, multiline=True)

    y -= 94
    draw_label(c, "Revision / Change Request Terms", left, y)
    y -= 16
    text_field(c, "revision_terms", left, y - 56, right - left, 60, multiline=True)

    y -= 78
    draw_label(c, "Client Responsibilities", left, y)
    y -= 16
    text_field(c, "client_responsibilities", left, y - 60, right - left, 64, multiline=True)

    c.setFont("Helvetica", 9)
    c.drawRightString(right, 18, "Page 2 of 3")
    c.showPage()

    # Page 3
    y = height - 45
    c.setFont("Helvetica-Bold", 14)
    c.drawString(left, y, "Acceptance and Signatures")
    y -= 30

    legal_lines = [
        "This proposal is accepted when signed by both parties and initial payment is received.",
        "Both parties agree to keep confidential information private.",
        "Post-launch bug support window: [customize per proposal].",
    ]
    y = draw_paragraph(c, legal_lines, left, y, line_height=14, size=10) - 12

    draw_label(c, "Client Name:", left, y)
    text_field(c, "accept_client_name", left + 95, y - 4, 210)
    draw_label(c, "Date:", left + 325, y)
    text_field(c, "accept_client_date", left + 360, y - 4, 165)

    y -= 34
    draw_label(c, "Client Signature:", left, y)
    text_field(c, "accept_client_signature", left + 110, y - 4, 415)

    y -= 34
    draw_label(c, "Freelancer Name:", left, y)
    text_field(c, "accept_freelancer_name", left + 110, y - 4, 195)
    draw_label(c, "Date:", left + 325, y)
    text_field(c, "accept_freelancer_date", left + 360, y - 4, 165)

    y -= 34
    draw_label(c, "Freelancer Signature:", left, y)
    text_field(c, "accept_freelancer_signature", left + 125, y - 4, 400)

    y -= 48
    draw_label(c, "Next Steps", left, y)
    y -= 16
    text_field(c, "next_steps", left, y - 72, right - left, 76, multiline=True)

    y -= 95
    draw_label(c, "Your Contact Info", left, y)
    y -= 16
    draw_label(c, "Email:", left, y)
    text_field(c, "your_email", left + 45, y - 4, 210)
    draw_label(c, "Phone:", left + 275, y)
    text_field(c, "your_phone", left + 320, y - 4, 205)

    y -= 30
    draw_label(c, "Website/Portfolio:", left, y)
    text_field(c, "your_website", left + 105, y - 4, 420)

    c.setFont("Helvetica", 9)
    c.drawRightString(right, 18, "Page 3 of 3")
    c.save()


if __name__ == "__main__":
    make_pdf("freelance/proposal-template-fillable.pdf")
    print("Created freelance/proposal-template-fillable.pdf")
