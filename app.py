"""
Active DMC Contact Report Agent — Web App
Run with: streamlit run app.py
"""

import os
import re
from datetime import datetime
import streamlit as st
import anthropic

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Active DMC · Contact Report Agent",
    page_icon="📋",
    layout="wide",
)

# ── Prompt template ───────────────────────────────────────────────────────────
PROMPT_TEMPLATE = """You are the Active DMC Contact Report Agent. Your job is to transform meeting data from Read AI (which includes transcripts, summaries, and action items) into a polished contact report following Active DMC's exact format.

Here is the meeting data you will be working with:

<meeting_data>
{meeting_data}
</meeting_data>

Your output must follow this exact structure:

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

Important rules to follow:

- Follow the Active DMC contact report format exactly as shown above
- Use "•" (bullet points) for main discussion topics and "o" (sub-bullets) for summary points under each topic
- Use numbered lists (1., 2.) for Next Steps owners, then "o" for individual action items
- Keep the tone professional, concise, and client-friendly
- Extract and synthesize information from the meeting data; do not copy the transcript word-for-word
- Organize discussion points by topic, not chronologically
- Preserve all specific details accurately: names, dates, deadlines, publications, clients, campaign details, deliverables, numbers, etc.
- For action items, carefully determine whether Active DMC or the client is responsible and place them under the correct owner
- If attendees, date, time, or location information is missing from the meeting data, use placeholders like "[Client Name]", "[Date]", "[Time]", or "[Location - Teams/Zoom/In person]"
- Always specify "UAE time" after the time
- Do not invent or assume information that is not present in the meeting data
- Write this as a professional follow-up email, not a meeting transcript or verbatim summary

Before writing your final contact report, use the scratchpad below to:
1. Identify the client name and company
2. List attendees from both sides
3. Extract meeting logistics (date, time, location)
4. Identify main discussion topics and organize key points under each
5. Extract action items and assign them to the correct owner (Active DMC or client)

<scratchpad>
[Your analysis and organization of the meeting data goes here]
</scratchpad>

Now write the complete contact report following the exact format specified above. Your final output should be the complete, polished contact report ready to send to the client - do not include the scratchpad in your final answer.

<contact_report>
"""


def generate_report(transcript: str) -> str:
    api_key = os.environ.get("ANTHROPIC_API_KEY") or st.session_state.get("api_key", "")
    if not api_key:
        st.error("No API key found. Set ANTHROPIC_API_KEY or enter it in the sidebar.")
        return ""

    client = anthropic.Anthropic(api_key=api_key)
    prompt = PROMPT_TEMPLATE.format(meeting_data=transcript)

    output_placeholder = st.empty()
    full_text = ""

    with client.messages.stream(
        model="claude-opus-4-7",
        max_tokens=20000,
        thinking={"type": "adaptive"},
        messages=[
            {
                "role": "user",
                "content": [{"type": "text", "text": prompt}],
            }
        ],
    ) as stream:
        for event in stream:
            # Only stream text deltas (skip thinking blocks)
            if (
                hasattr(event, "type")
                and event.type == "content_block_delta"
                and hasattr(event, "delta")
                and getattr(event.delta, "type", "") == "text_delta"
            ):
                full_text += event.delta.text
                output_placeholder.markdown(
                    f"```\n{full_text}\n```",
                    unsafe_allow_html=False,
                )

    # Strip closing </contact_report> tag if present
    full_text = re.sub(r"\s*</contact_report>\s*$", "", full_text).strip()
    output_placeholder.empty()
    return full_text


# ── Sidebar — API key ─────────────────────────────────────────────────────────
with st.sidebar:
    st.image(
        "https://www.activedmc.com/wp-content/uploads/2022/06/Active-DMC-Logo.png",
        use_column_width=True,
    ) if False else st.markdown("## 📋 Active DMC")  # logo placeholder — replace URL if needed

    st.markdown("### Contact Report Agent")
    st.markdown("Paste a meeting transcript and generate a polished client contact report.")
    st.divider()

    env_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if env_key:
        st.success("API key loaded from environment.", icon="✅")
    else:
        manual_key = st.text_input(
            "Anthropic API Key",
            type="password",
            placeholder="sk-ant-...",
            help="Or set the ANTHROPIC_API_KEY environment variable.",
        )
        if manual_key:
            st.session_state["api_key"] = manual_key
            st.success("API key saved for this session.", icon="✅")

    st.divider()
    st.caption("Model: claude-opus-4-7 · Adaptive thinking")


# ── Main UI ───────────────────────────────────────────────────────────────────
st.title("Contact Report Agent")
st.markdown("Paste your Read AI transcript or meeting notes below, then click **Generate Report**.")

transcript = st.text_area(
    label="Meeting transcript / notes",
    placeholder=(
        "Paste the full Read AI report here — transcript, summary, action items, "
        "attendees, topics, or any combination."
    ),
    height=320,
    label_visibility="collapsed",
)

col1, col2 = st.columns([1, 5])
with col1:
    generate_clicked = st.button("Generate Report", type="primary", use_container_width=True)
with col2:
    if st.session_state.get("report"):
        st.caption("✅ Report ready — copy from the box below or use the Download button.")

st.divider()

# ── Generate ──────────────────────────────────────────────────────────────────
if generate_clicked:
    if not transcript.strip():
        st.warning("Please paste a transcript before generating.")
    else:
        with st.spinner("Generating contact report…"):
            report = generate_report(transcript)
        if report:
            st.session_state["report"] = report
            st.session_state["report_time"] = datetime.now().strftime("%Y-%m-%d %H:%M")

# ── Output ────────────────────────────────────────────────────────────────────
if st.session_state.get("report"):
    report = st.session_state["report"]
    ts     = st.session_state.get("report_time", "")

    st.subheader("Contact Report")
    if ts:
        st.caption(f"Generated {ts}")

    st.text_area(
        label="report_output",
        value=report,
        height=520,
        label_visibility="collapsed",
    )

    filename = f"ADMC_ContactReport_{datetime.now().strftime('%Y-%m-%d_%H%M')}.txt"
    st.download_button(
        label="⬇ Download .txt",
        data=report,
        file_name=filename,
        mime="text/plain",
    )

    if st.button("Clear & Start Over"):
        st.session_state.pop("report", None)
        st.session_state.pop("report_time", None)
        st.rerun()
