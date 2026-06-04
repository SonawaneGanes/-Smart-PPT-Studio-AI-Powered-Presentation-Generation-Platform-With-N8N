import streamlit as st
import subprocess
import os
import sys
import traceback

# ---------------- PAGE CONFIG ----------------

st.set_page_config(
    page_title="Smart PPT Studio",
    page_icon="🚀",
    layout="wide"
)

# ---------------- CSS ----------------

st.markdown("""
<style>

.main {
    background: #0f172a;
}

h1, h2, h3, p {
    color: white;
}

textarea {
    border-radius: 12px !important;
}

.stButton > button {

    width: 100%;

    background: linear-gradient(
        90deg,
        #8b5cf6,
        #3b82f6
    );

    color: white;

    border: none;

    border-radius: 14px;

    padding: 14px;

    font-size: 18px;

    font-weight: bold;
}

.card {

    background: rgba(255,255,255,0.05);

    padding: 30px;

    border-radius: 20px;

    border: 1px solid rgba(255,255,255,0.08);
}

</style>
""", unsafe_allow_html=True)

# ---------------- HEADER ----------------

st.markdown("""
<h1 style='text-align:center;'>
🚀 Smart PPT Studio
</h1>
""", unsafe_allow_html=True)

st.markdown("""
<p style='text-align:center;color:#94a3b8;font-size:18px;'>
Generate professional PowerPoint presentations instantly
</p>
""", unsafe_allow_html=True)

# ---------------- CARD ----------------

st.markdown("<div class='card'>", unsafe_allow_html=True)

prompt = st.text_area(
    "Presentation Topic",
    placeholder="Example: Artificial Intelligence in Healthcare",
    height=200
)

generate_btn = st.button(
    "✨ Generate Presentation"
)

# ---------------- GENERATION ----------------

if generate_btn:

    if not prompt.strip():

        st.warning(
            "Please enter a topic."
        )

    else:

        try:

            # ---------------- LOAD TEMPLATE ----------------

            with open(
                "app1.py",
                "r",
                encoding="utf-8"
            ) as file:

                template_code = file.read()

            # ---------------- DETECT THEME ----------------

            theme = "blue"

            lower_prompt = prompt.lower()

            if "finance" in lower_prompt:
                theme = "green"

            elif "health" in lower_prompt:
                theme = "red"

            elif "cricket" in lower_prompt:
                theme = "purple"

            elif "ai" in lower_prompt:
                theme = "blue"

            # ---------------- SAFE STRING ----------------

            safe_prompt = repr(prompt)

            # ---------------- REPLACE VARIABLES ----------------

            generated_code = template_code

            generated_code = generated_code.replace(
                "{{TOPIC}}",
                safe_prompt
            )

            generated_code = generated_code.replace(
                "{{THEME}}",
                theme
            )

            # ---------------- DELETE OLD FILES ----------------

            if os.path.exists("generated_presentation.py"):
                os.remove("generated_presentation.py")

            if os.path.exists("modern_presentation.pptx"):
                os.remove("modern_presentation.pptx")

            # ---------------- WRITE NEW FILE ----------------

            with open(
                "generated_presentation.py",
                "w",
                encoding="utf-8"
            ) as file:

                file.write(generated_code)

            # ---------------- EXECUTE ----------------

            process = subprocess.run(
                [
                    sys.executable,
                    "generated_presentation.py"
                ],
                capture_output=True,
                text=True
            )

            # ---------------- DEBUG ----------------

            print(process.stdout)
            print(process.stderr)

            # ---------------- HANDLE ERRORS ----------------

            if process.returncode != 0:

                st.error("PowerPoint generation failed.")

                st.code(process.stderr)

            else:

                ppt_file = "modern_presentation.pptx"

                if os.path.exists(ppt_file):

                    st.success(
                        f"Presentation generated successfully on topic:\n\n{prompt}"
                    )

                    with open(ppt_file, "rb") as file:

                        st.download_button(
                            label="⬇ Download Presentation",
                            data=file,
                            file_name="Smart_Presentation.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )

                else:

                    st.error("Presentation file not found.")

        except Exception:

            st.error(
                "Unexpected system error occurred."
            )

            st.code(
                traceback.format_exc()
            )

st.markdown("</div>", unsafe_allow_html=True)
