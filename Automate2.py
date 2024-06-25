import os
import openai
import pandas as pd
import time
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="CERF Level Sheet")

def main():
    st.header("Ask your Sheet")

    # Set the API key through an environment variable
    api_key = st.text_input("Enter your OpenAI API Key", type="password")
    if api_key:
        os.environ["OPENAI_API_KEY"] = api_key
        client = openai.OpenAI(api_key=api_key)

    file_path = st.file_uploader("Upload your Excel Sheet", type="xlsx")

    attribute = st.text_input("Enter the attribute you want to retrieve (e.g., 'CEFR level', 'part of speech')")

    if file_path is not None and api_key and attribute:
        df = pd.read_excel(file_path)
        words = df['Table 1'].dropna().tolist()[0:800]

        progress_bar = st.progress(0)
        progress_text = st.empty()

        def get_word_attribute(word, attribute):
            retries = 0
            max_retries = 5
            backoff_factor = 1.5

            while retries < max_retries:
                try:
                    response = client.chat.completions.create(
                        messages=[
                            {"role": "system", "content": f"You are a helpful assistant knowledgeable about {attribute}. Only give a one-word answer which is the {attribute} of the word asked."},
                            {"role": "user", "content": f"What is the {attribute} of the word '{word}'?"}
                        ],
                        model="gpt-3.5-turbo",
                        max_tokens=1000
                    )
                    return response.choices[0].message.content.strip()
                except Exception as e:
                    if "Rate limit exceeded" in str(e):
                        wait = (backoff_factor ** retries) * 2
                        time.sleep(wait)
                        retries += 1
                    else:
                        st.error(f"An error occurred: {e}")
                        break

            return "Rate limit exceeded, try again later."

        word_attributes = []

        for i, word in enumerate(words):
            try:
                word_attribute = get_word_attribute(word, attribute)
                word_attributes.append((word, word_attribute))
                time.sleep(1)  # Sleep to avoid hitting rate limits

                # Update progress bar and text
                progress_percent = min((i + 1) / len(words), 1.0)  # Properly calculate progress as a fraction
                progress_bar.progress(progress_percent)
                progress_text.text(f"Processing word {i + 1} of {len(words)}")

            except Exception as e:
                st.error(f"An error occurred for word '{word}': {e}")
                word_attributes.append((word, 'Error'))

        progress_bar.empty()
        progress_text.empty()

        attributes_df = pd.DataFrame(word_attributes, columns=['Word', attribute])
        output_file_path = 'word_attributes.xlsx'
        attributes_df.to_excel(output_file_path, index=False)

        st.success(f'{attribute.capitalize()}s saved to {output_file_path}')

        buffer = BytesIO()
        attributes_df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            label=f"Download {attribute.capitalize()} Data as Excel",
            data=buffer,
            file_name="word_attributes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if __name__ == "__main__":
    main()