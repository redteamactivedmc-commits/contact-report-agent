#!/usr/bin/env python3
"""
Active DMC Contact Report Agent
Reads a meeting transcript and produces a formatted contact report.

Usage:
    python contact_report_agent.py transcript.txt
    python contact_report_agent.py transcript.txt --output report.txt
    cat transcript.txt | python contact_report_agent.py
    python contact_report_agent.py   # interactive paste mode
"""

import sys
import os
import argparse
import anthropic

SYSTEM_PROMPT = """You are the Active DMC Contact Report Agent.

Your job is to take a Read AI meeting report, transcript, summary, and action items, then produce a polished contact report in Active DMC's exact format.

Use this structure every time:

Hi [Client Name],

Thanks for your time – following our discussion please see the contact report below.

Present on behalf of [Client Company]:                       [Client attendees]

Present on behalf of Active DMC:                       [Active DMC attendees]

Date:                                                                [Date]

Time:                                                               [Time] UAE time

Location:                                                          [Teams/Zoom/In person]

Summary of Discussion

•\t[Topic title]
o\t[Summary point]
o\t[Summary point]

•\t[Topic title]
o\t[Summary point]
o\t[Summary point]

Next Steps

1.\tActive DMC
o\t[Action point owned by Active DMC]
o\t[Action point owned by Active DMC]

2.\t[Client Name/Client Company]
o\t[Action point owned by client]
o\t[Action point owned by client]

Please let us know if we missed anything or if you have any questions.

Rules:
- Follow the Active DMC contact report format exactly.
- Use "•" for main discussion points and "o" for sub-points.
- Keep the tone professional, concise, and client-friendly.
- Extract action items from the transcript and place them under the correct owner.
- Do not copy the transcript word for word.
- Do not invent missing information.
- If attendees, date, time, or location are missing, use placeholders like [Date], [Time], [Location].
- Preserve all names, deadlines, publications, clients, campaign details, and deliverables accurately.
- Write like a post-meeting follow-up email, not a meeting transcript.
- Active DMC staff are identified by their @activedmc.com email addresses or by context clues indicating they represent Active DMC."""


def read_transcript_from_file(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def read_transcript_from_stdin() -> str:
    print("Reading transcript from stdin...", file=sys.stderr)
    return sys.stdin.read()


def read_transcript_interactive() -> str:
    print("Paste the meeting transcript below.")
    print("When done, press Enter then Ctrl+D (Mac/Linux) or Ctrl+Z then Enter (Windows).")
    print("-" * 60)
    lines = []
    try:
        while True:
            line = input()
            lines.append(line)
    except EOFError:
        pass
    return "\n".join(lines)


def generate_contact_report(transcript: str, save_path: str | None = None) -> str:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("Error: ANTHROPIC_API_KEY environment variable not set.", file=sys.stderr)
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    print("Generating contact report...", file=sys.stderr)

    report_text = ""
    with client.messages.stream(
        model="claude-opus-4-7",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": f"Please generate a contact report from the following meeting transcript:\n\n{transcript}",
            }
        ],
    ) as stream:
        for text in stream.text_stream:
            print(text, end="", flush=True)
            report_text += text

    print()  # newline after streamed output

    if save_path:
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(report_text)
        print(f"\nReport saved to: {save_path}", file=sys.stderr)

    return report_text


def main():
    parser = argparse.ArgumentParser(
        description="Generate an Active DMC contact report from a meeting transcript."
    )
    parser.add_argument(
        "transcript_file",
        nargs="?",
        help="Path to the transcript file. Reads from stdin if omitted.",
    )
    parser.add_argument(
        "--output", "-o",
        metavar="OUTPUT_FILE",
        help="Save the generated report to this file.",
    )
    args = parser.parse_args()

    if args.transcript_file:
        transcript = read_transcript_from_file(args.transcript_file)
    elif not sys.stdin.isatty():
        transcript = read_transcript_from_stdin()
    else:
        transcript = read_transcript_interactive()

    if not transcript.strip():
        print("Error: No transcript content provided.", file=sys.stderr)
        sys.exit(1)

    generate_contact_report(transcript, save_path=args.output)


if __name__ == "__main__":
    main()
