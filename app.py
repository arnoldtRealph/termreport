import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from io import BytesIO
from docx import Document
from docx.shared import Inches

# Streamlit Config
st.set_page_config(page_title="📊 Learner Performance Dashboard", layout="wide")

# Custom Styling
st.markdown("""
    <style>
        .stApp {
            background-color: #F4F6F9;
            color: #333333;
            font-family: 'Arial', sans-serif;
        }
        h1, h2, h3 {
            color: #003366;
        }
        .custom-title {
            font-size: 30px;
            color: #003366;
        }
        .welcome-text {
            font-size: 20px;
            color: #4CAF50;
            text-align: center;
            margin-bottom: 20px;
        }
        .instructions {
            font-size: 16px;
            color: #333333;
            text-align: center;
            margin-bottom: 20px;
        }
        .footer {
            text-align: center;
            font-size: 12px;
            color: #003366;
            margin-top: 20px;
        }
    </style>
""", unsafe_allow_html=True)

# Sidebar Configuration
st.sidebar.title("📂 Options")
uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

sections = {
    "Average per Learner": "overview",
    "Question Analysis": "charts",
    "Insights and Recommendations": "recommendations",
    "Download Full Report": "download"
}

for section_name, section_id in sections.items():
    st.sidebar.markdown(f'<a href="#{section_id}" class="custom-button">{section_name}</a>', unsafe_allow_html=True)

# Main content with Welcome Message and Instructions
st.markdown("<h1 class='custom-title' style='text-align: center;'>Saul Damon High School</h1>", unsafe_allow_html=True)
st.markdown("""
    <div class='welcome-text'>
        Welkom julle. Gebruik hierdie handige tool wat ek ontwerp het om handige insig te kry vanuit leerders se toetse en take.
    </div>
    <div class='instructions'>
        To get started, upload your Excel mark sheet in the sidebar. Ensure it includes learner names, 
        question scores (e.g., Q1, Q2), and optionally a 'Total' column. Let’s transform numbers into actionable insights!
    </div>
""", unsafe_allow_html=True)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if 'NO' in df.columns:
        df = df.drop(columns=['NO'])

    name_col = None
    total_col = None
    for col in df.columns:
        if "name" in col.lower():
            name_col = col
        if "total" in col.lower():
            total_col = col

    numeric_cols = df.select_dtypes(include=['number']).columns
    question_cols = [col for col in numeric_cols if 'q' in col.lower() or 'question' in col.lower()]

    if total_col:
        df['Percentage'] = (df[total_col] / df[total_col].max()) * 100
    elif question_cols:
        df['Total'] = df[question_cols].sum(axis=1)
        df['Percentage'] = (df['Total'] / df['Total'].max()) * 100
        total_col = 'Total'

    # Define function to save the chart to a buffer
    def save_chart_to_buffer(fig):
        buf = BytesIO()
        fig.savefig(buf, format="png")
        buf.seek(0)
        return buf

    # Section: Results Overview
    if st.markdown(f'<a name="overview"></a>', unsafe_allow_html=True):
        st.subheader("Results DASHBOARD")
        st.write("Upload your class mark sheet to generate insightful analysis.")
        st.subheader("📊 Average Percentage per Learner")
        if name_col and total_col:
            avg_fig, ax1 = plt.subplots(figsize=(8, max(5, len(df) * 0.3)))
            bars = ax1.barh(df[name_col], df['Percentage'], color=sns.color_palette("Blues_d", len(df)))
            for bar in bars:
                bar.set_edgecolor('#003366')
                bar.set_linewidth(1)
            ax1.set_title("Percentage per Learner", color='#003366', pad=20)
            ax1.set_xlabel("Percentage (%)")
            ax1.set_ylabel("Learner")
            ax1.invert_yaxis()
            ax1.grid(True, axis='x', linestyle='--', alpha=0.7)
            plt.tight_layout()
            st.pyplot(avg_fig)

    # Section: Question Analysis
    if st.markdown(f'<a name="charts"></a>', unsafe_allow_html=True):
        st.subheader("📊 Question Analysis")
        if question_cols:
            st.subheader("Average Marks per Question")
            q_avg_fig, ax_q = plt.subplots(figsize=(8, 5))
            question_means = df[question_cols].mean()
            sns.barplot(x=question_means.index, y=question_means.values, ax=ax_q, palette="Greens_d")
            ax_q.set_title("Average Marks per Question", color='#003366')
            ax_q.set_xlabel("Question")
            ax_q.set_ylabel("Average Mark")
            ax_q.set_xticklabels(ax_q.get_xticklabels(), rotation=45, ha='right')
            st.pyplot(q_avg_fig)

            st.subheader("Distribution per Question")
            for question in question_cols:
                pie_fig, ax2 = plt.subplots(figsize=(6, 6))
                question_data = df[question].value_counts()
                total = question_data.sum()
                percentages = [(count / total * 100) for count in question_data]
                wedges, texts, autotexts = ax2.pie(
                    question_data,
                    startangle=90,
                    colors=sns.color_palette("Set2", len(question_data)),
                    autopct='%1.1f%%',
                    pctdistance=0.85,
                    wedgeprops={'edgecolor': 'white', 'linewidth': 1}
                )
                ax2.axis('equal')
                ax2.set_title(f"Distribution of Marks for {question}", color='#003366', pad=20)
                for autotext in autotexts:
                    autotext.set_color('white')
                    autotext.set_fontsize(10)
                    autotext.set_weight('bold')
                ax2.legend(
                    labels=[f"{mark}: {pct:.1f}%" for mark, pct in zip(question_data.index, percentages)],
                    title="Marks (Percentage)",
                    loc="center left",
                    bbox_to_anchor=(1, 0, 0.5, 1)
                )
                plt.tight_layout()
                st.pyplot(pie_fig)

    # Section: Insights and Recommendations
    if st.markdown(f'<a name="recommendations"></a>', unsafe_allow_html=True):
        st.subheader("📝 Insights and Recommendations")
        insights = []
        recommendations = []

        if name_col and total_col:
            avg_percentage = df['Percentage'].mean()
            insights.append(f"The class average percentage is {avg_percentage:.2f}%.")
            failing_learners = df[df['Percentage'] < 30][name_col].tolist()
            if failing_learners:
                insights.append(f"Learners failing (<30%): {', '.join(failing_learners)}.")
                recommendations.append("Organize targeted remedial sessions for failing learners, focusing on foundational skills and consistent practice.")
            distinction_learners = df[df['Percentage'] >= 80][name_col].tolist()
            if distinction_learners:
                insights.append(f"Learners with distinction (≥80%): {', '.join(distinction_learners)}.")
                recommendations.append("Encourage distinction learners with advanced assignments or leadership roles in study groups to deepen their understanding.")
            if question_cols:
                question_means = df[question_cols].mean()
                avg_question_mean = question_means.mean()
                weak_questions = question_means[question_means < avg_question_mean * 0.7].index.tolist()
                if weak_questions:
                    weak_avg = question_means[weak_questions].mean()
                    insights.append(f"Weak questions (below 70% of mean): {', '.join(weak_questions)} with an average of {weak_avg:.1f}.")
                    recommendations.append(f"Focus revision on {', '.join(weak_questions)}.")
        
        st.table([{'Insight': item} for item in insights])
        st.table([{'Recommendation': item} for item in recommendations])

    # Section: Download Full Report (Word)
    if st.markdown(f'<a name="download"></a>', unsafe_allow_html=True):
        st.subheader("📥 Download Full Report (Word Document)")
        if st.button("Download Full Report"):
            # Create a Word Document
            doc = Document()

            # Title
            doc.add_heading('Learner Performance Report', 0)

            # Add Average Percentage per Learner chart
            if name_col and total_col:
                avg_fig, ax1 = plt.subplots(figsize=(8, max(5, len(df) * 0.3)))
                bars = ax1.barh(df[name_col], df['Percentage'], color=sns.color_palette("Blues_d", len(df)))
                for bar in bars:
                    bar.set_edgecolor('#003366')
                    bar.set_linewidth(1)
                ax1.set_title("Percentage per Learner", color='#003366', pad=20)
                ax1.set_xlabel("Percentage (%)")
                ax1.set_ylabel("Learner")
                ax1.invert_yaxis()
                ax1.grid(True, axis='x', linestyle='--', alpha=0.7)
                plt.tight_layout()

                # Save to buffer and add to Word
                img_stream = save_chart_to_buffer(avg_fig)
                doc.add_picture(img_stream, width=Inches(5.0))

            # Add Insights and Recommendations
            doc.add_heading('Insights and Recommendations', level=1)
            doc.add_paragraph("\n".join(insights), style='BodyText')
            doc.add_paragraph("\n".join(recommendations), style='BodyText')

            # Add the Question Analysis charts (Bar and Pie)
            if question_cols:
                # Add Average Marks per Question chart
                q_avg_fig, ax_q = plt.subplots(figsize=(8, 5))
                question_means = df[question_cols].mean()
                sns.barplot(x=question_means.index, y=question_means.values, ax=ax_q, palette="Greens_d")
                ax_q.set_title("Average Marks per Question", color='#003366')
                ax_q.set_xlabel("Question")
                ax_q.set_ylabel("Average Mark")
                ax_q.set_xticklabels(ax_q.get_xticklabels(), rotation=45, ha='right')
                st.pyplot(q_avg_fig)

                # Save to buffer and add to Word
                q_avg_stream = save_chart_to_buffer(q_avg_fig)
                doc.add_picture(q_avg_stream, width=Inches(5.0))

                # Add Distribution per Question Pie chart
                for question in question_cols:
                    pie_fig, ax2 = plt.subplots(figsize=(6, 6))
                    question_data = df[question].value_counts()
                    total = question_data.sum()
                    percentages = [(count / total * 100) for count in question_data]
                    wedges, texts, autotexts = ax2.pie(
                        question_data,
                        startangle=90,
                        colors=sns.color_palette("Set2", len(question_data)),
                        autopct='%1.1f%%',
                        pctdistance=0.85,
                        wedgeprops={'edgecolor': 'white', 'linewidth': 1}
                    )
                    ax2.axis('equal')
                    ax2.set_title(f"Distribution of Marks for {question}", color='#003366', pad=20)
                    for autotext in autotexts:
                        autotext.set_color('white')
                        autotext.set_fontsize(10)
                        autotext.set_weight('bold')
                    ax2.legend(
                        labels=[f"{mark}: {pct:.1f}%" for mark, pct in zip(question_data.index, percentages)],
                        title="Marks (Percentage)",
                        loc="center left",
                        bbox_to_anchor=(1, 0, 0.5, 1)
                    )
                    plt.tight_layout()

                    # Save to buffer and add to Word
                    pie_stream = save_chart_to_buffer(pie_fig)
                    doc.add_picture(pie_stream, width=Inches(5.0))

            # Save to memory buffer instead of file
            doc_stream = BytesIO()
            doc.save(doc_stream)
            doc_stream.seek(0)

            # Provide download link for Word file
            st.download_button(
                label="Download Full Report (Word Document)",
                data=doc_stream,
                file_name="learner_performance_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# Footer
st.markdown(f'<div class="footer">Designed by Mr AR Visagie</div>', unsafe_allow_html=True)
