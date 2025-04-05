import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from io import BytesIO
from docx import Document
from docx.shared import Inches

# Streamlit Config
st.set_page_config(page_title="📊 Learner Performance Dashboard", layout="wide")

# Custom Styling with Explicit Colors
st.markdown("""
    <style>
        .stApp {
            background-color: #F4F6F9;  /* Light gray-blue background */
            color: #333333;            /* Dark gray text */
            font-family: Arial, sans-serif;
        }
        .custom-title {
            font-size: 30px;
            color: #003366;            /* Deep blue */
        }
        .welcome-text {
            font-size: 20px;
            color: #4CAF50;            /* Green for welcome text */
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
            color: #003366;            /* Deep blue */
            margin-top: 20px;
        }
        .sidebar .sidebar-content {
            background-color: #E8ECEF;  /* Light gray sidebar */
            padding: 10px;
        }
        .sidebar h3 {
            font-size: 18px;
            color: #003366;            /* Deep blue */
        }
        h2, h3 {
            color: #003366;            /* Deep blue for subheaders */
        }
        .stButton>button {
            background-color: #003366; /* Deep blue buttons */
            color: white;
        }
    </style>
""", unsafe_allow_html=True)

# Sidebar Configuration
st.sidebar.title("📊 Dashboard Options")
st.sidebar.markdown("### File Upload")
uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

st.sidebar.markdown("### Navigation")
sections = {
    "Learner Overview": "overview",
    "Question Analysis": "charts",
    "Insights and Recommendations": "recommendations",
    "Download Full Report": "download"
}
for section_name, section_id in sections.items():
    st.sidebar.markdown(f'<a href="#{section_id}" style="text-decoration: none; color: #003366;">{section_name}</a>', unsafe_allow_html=True)

st.sidebar.markdown("### Chart Settings")
chart_options = ["Vertical Bar", "Scatter Plot", "Box Plot"]
selected_chart = st.sidebar.selectbox("Select chart type for Total Marks per Learner", chart_options, index=0)

# Main content with Welcome Message and Instructions in Markdown
st.markdown("<h1 class='custom-title' style='text-align: center;'>Saul Damon High School</h1>", unsafe_allow_html=True)
st.markdown("""
<div class='welcome-text'>
**Welkom julle.** Gebruik hierdie tool wat ek ontwerp het om handige insig te kry vanuit leerders se toetse en take.
</div>

<div class='instructions'>
1. Gebruik die excel templaat wat ek aan julle voorsien het om die punte per vraag analise te doen.  
2. Upload dan die excel file in die sidebar hier langsaan.  
3. Download dan die Word dokument en stoor in die masterfile.
</div>
""", unsafe_allow_html=True)

if uploaded_file:
    # Read the Excel file without assuming a header initially
    df_raw = pd.read_excel(uploaded_file, header=None)

    # Find the row containing "NAME OF LEARNER"
    header_row = None
    for i, row in df_raw.iterrows():
        if any("name of learner" in str(cell).lower() for cell in row):
            header_row = i
            break

    if header_row is not None:
        # Use the identified row as the header and skip rows above it
        df = pd.read_excel(uploaded_file, skiprows=header_row)
        df.columns = df.columns.str.strip()  # Clean column names

        # Find the "NAME OF LEARNER" column
        name_col = None
        for col in df.columns:
            if "name of learner" in col.lower():
                name_col = col
                break

        if name_col:
            # Get question columns (numeric columns after name_col), handle NaN gracefully
            name_col_idx = df.columns.get_loc(name_col)
            question_cols = [col for col in df.columns[name_col_idx + 1:] if df[col].dtype in ['int64', 'float64']]
            
            if question_cols:
                # Fill NaN with 0 for calculations, assuming unmarked questions are 0
                df[question_cols] = df[question_cols].fillna(0)
                df['Total'] = df[question_cols].sum(axis=1)
                # Avoid division by zero if all totals are 0
                max_total = df['Total'].max()
                df['Percentage'] = (df['Total'] / max_total * 100) if max_total > 0 else 0
                total_col = 'Total'
            else:
                st.error("No numeric question columns found after 'NAME OF LEARNER'.")
                st.stop()
        else:
            st.error("Could not find 'NAME OF LEARNER' column after setting headers.")
            st.stop()
    else:
        st.error("Could not locate a row containing 'NAME OF LEARNER' in the file.")
        st.stop()

    # Define function to save the chart to a buffer
    def save_chart_to_buffer(fig):
        buf = BytesIO()
        fig.savefig(buf, format="png")
        buf.seek(0)
        return buf

    # Section: Learner Overview
    st.markdown(f'<a name="overview"></a>', unsafe_allow_html=True)
    st.subheader("Results DASHBOARD")
    st.write("Upload your class mark sheet to generate insightful analysis.")
    
    # Average Percentage per Learner
    st.subheader("📊 Average Percentage per Learner")
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

    # Total Marks per Learner
    st.subheader("📊 Total Marks per Learner")
    if selected_chart == "Vertical Bar":
        total_fig, ax1 = plt.subplots(figsize=(10, 6))
        sns.barplot(x=df[name_col], y=df['Total'], ax=ax1, palette="Blues_d")
        ax1.set_title("Total Marks per Learner", color='#003366', pad=20)
        ax1.set_xlabel("Learner")
        ax1.set_ylabel("Total Marks")
        ax1.set_xticklabels(ax1.get_xticklabels(), rotation=90, ha='center')
        ax1.grid(True, axis='y', linestyle='--', alpha=0.7)
        plt.tight_layout()
        st.pyplot(total_fig)

    elif selected_chart == "Scatter Plot":
        total_fig, ax1 = plt.subplots(figsize=(10, 6))
        ax1.scatter(df[name_col], df['Total'], color='#003366', edgecolor='black', s=100)
        ax1.set_title("Total Marks per Learner", color='#003366', pad=20)
        ax1.set_xlabel("Learner")
        ax1.set_ylabel("Total Marks")
        ax1.set_xticklabels(df[name_col], rotation=60, ha='right')
        ax1.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()
        st.pyplot(total_fig)

    elif selected_chart == "Box Plot":
        total_fig, ax1 = plt.subplots(figsize=(10, 6))
        sns.boxplot(y=df['Total'], ax=ax1, color='#003366', boxprops=dict(edgecolor='black'))
        ax1.set_title("Distribution of Total Marks", color='#003366', pad=20)
        ax1.set_ylabel("Total Marks")
        ax1.grid(True, axis='y', linestyle='--', alpha=0.7)
        plt.tight_layout()
        st.pyplot(total_fig)

    # Section: Question Analysis
    st.markdown(f'<a name="charts"></a>', unsafe_allow_html=True)
    st.subheader("📊 Question Analysis")
    active_questions = [q for q in question_cols if df[q].sum() > 0]
    
    if active_questions:
        st.subheader("Average Marks per Question")
        q_avg_fig, ax_q = plt.subplots(figsize=(8, 5))
        active_means = df[active_questions].mean()
        sns.barplot(x=active_means.index, y=active_means.values, ax=ax_q, palette="Greens_d")
        ax_q.set_title("Average Marks per Question", color='#003366')
        ax_q.set_xlabel("Question")
        ax_q.set_ylabel("Average Mark")
        ax_q.set_xticklabels(ax_q.get_xticklabels(), rotation=45, ha='right')
        st.pyplot(q_avg_fig)

        st.subheader("Distribution per Question")
        for question in active_questions:
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
    else:
        st.write("No questions have marks entered yet.")

    # Section: Insights and Recommendations
    st.markdown(f'<a name="recommendations"></a>', unsafe_allow_html=True)
    st.subheader("📝 Insights and Recommendations")
    
    insights = []
    recommendations = []

    avg_percentage = df['Percentage'].mean()
    insights.append(f"The class average percentage is **{avg_percentage:.2f}%**.")
    failing_learners = df[df['Percentage'] < 30][name_col].tolist()
    if failing_learners:
        insights.append(f"Learners failing (<30%): **{', '.join(failing_learners)}**.")
        recommendations.append("Organize targeted remedial sessions for failing learners, focusing on foundational skills and consistent practice.")
    distinction_learners = df[df['Percentage'] >= 80][name_col].tolist()
    if distinction_learners:
        insights.append(f"Learners with distinction (≥80%): **{', '.join(distinction_learners)}**.")
        recommendations.append("Encourage distinction learners with advanced assignments or leadership roles in study groups to deepen their understanding.")
    
    if active_questions:
        active_means = df[active_questions].mean()
        avg_question_mean = active_means.mean()
        weak_questions = active_means[active_means < avg_question_mean * 0.7].index.tolist()
        if weak_questions:
            weak_avg = active_means[weak_questions].mean()
            insights.append(f"Weak questions (below 70% of mean): **{', '.join(weak_questions)}** with an average of **{weak_avg:.1f}**.")
            recommendations.append(f"Focus revision on **{', '.join(weak_questions)}**.")
    else:
        insights.append("No questions have marks entered yet.")

    st.markdown("### Insights")
    for insight in insights:
        st.markdown(f"- {insight}")

    st.markdown("### Recommendations")
    for recommendation in recommendations:
        st.markdown(f"- {recommendation}")

    # Section: Download Full Report (Word)
    st.markdown(f'<a name="download"></a>', unsafe_allow_html=True)
    st.subheader("📥 Download Full Report (Word Document)")
    if st.button("Download Full Report"):
        doc = Document()
        doc.add_heading('Learner Performance Report', 0)

        # Add Average Percentage per Learner chart
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
        img_stream = save_chart_to_buffer(avg_fig)
        doc.add_picture(img_stream, width=Inches(5.0))

        # Add Total Marks per Learner chart (using selected chart type)
        if selected_chart == "Vertical Bar":
            total_fig, ax1 = plt.subplots(figsize=(10, 6))
            sns.barplot(x=df[name_col], y=df['Total'], ax=ax1, palette="Blues_d")
            ax1.set_title("Total Marks per Learner", color='#003366', pad=20)
            ax1.set_xlabel("Learner")
            ax1.set_ylabel("Total Marks")
            ax1.set_xticklabels(ax1.get_xticklabels(), rotation=90, ha='center')
            ax1.grid(True, axis='y', linestyle='--', alpha=0.7)
            plt.tight_layout()
            total_stream = save_chart_to_buffer(total_fig)
            doc.add_picture(total_stream, width=Inches(5.0))

        elif selected_chart == "Scatter Plot":
            total_fig, ax1 = plt.subplots(figsize=(10, 6))
            ax1.scatter(df[name_col], df['Total'], color='#003366', edgecolor='black', s=100)
            ax1.set_title("Total Marks per Learner", color='#003366', pad=20)
            ax1.set_xlabel("Learner")
            ax1.set_ylabel("Total Marks")
            ax1.set_xticklabels(df[name_col], rotation=90, ha='center')
            ax1.grid(True, linestyle='--', alpha=0.7)
            plt.tight_layout()
            total_stream = save_chart_to_buffer(total_fig)
            doc.add_picture(total_stream, width=Inches(5.0))

        elif selected_chart == "Box Plot":
            total_fig, ax1 = plt.subplots(figsize=(10, 6))
            sns.boxplot(y=df['Total'], ax=ax1, color='#003366', boxprops=dict(edgecolor='black'))
            ax1.set_title("Distribution of Total Marks", color='#003366', pad=20)
            ax1.set_ylabel("Total Marks")
            ax1.grid(True, axis='y', linestyle='--', alpha=0.7)
            plt.tight_layout()
            total_stream = save_chart_to_buffer(total_fig)
            doc.add_picture(total_stream, width=Inches(5.0))

        # Add Insights and Recommendations
        doc.add_heading('Insights and Recommendations', level=1)
        doc.add_paragraph("\n".join([insight.replace('**', '') for insight in insights]), style='BodyText')
        doc.add_paragraph("\n".join([rec.replace('**', '') for rec in recommendations]), style='BodyText')

        # Add Question Analysis charts for active questions
        if active_questions:
            q_avg_fig, ax_q = plt.subplots(figsize=(8, 5))
            sns.barplot(x=active_means.index, y=active_means.values, ax=ax_q, palette="Greens_d")
            ax_q.set_title("Average Marks per Question", color='#003366')
            ax_q.set_xlabel("Question")
            ax_q.set_ylabel("Average Mark")
            ax_q.set_xticklabels(ax_q.get_xticklabels(), rotation=45, ha='right')
            q_avg_stream = save_chart_to_buffer(q_avg_fig)
            doc.add_picture(q_avg_stream, width=Inches(5.0))

            for question in active_questions:
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
                pie_stream = save_chart_to_buffer(pie_fig)
                doc.add_picture(pie_stream, width=Inches(5.0))

        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        st.download_button(
            label="Download Full Report (Word Document)",
            data=doc_stream,
            file_name="learner_performance_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# Footer
st.markdown(f'<div class="footer">Designed by Mr AR Visagie</div>', unsafe_allow_html=True)