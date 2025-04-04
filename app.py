import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from fpdf import FPDF
from io import BytesIO
import os

# Streamlit Config
st.set_page_config(page_title="📊 Learner Performance Dashboard", layout="wide")

# Custom Styling
st.markdown("""
    <style>
        .stApp {
            background-color: #F4F6F9;  /* Light background */
            color: #333333;  /* Dark text for readability */
            font-family: 'Arial', sans-serif;
        }
        h1, h2, h3 {
            color: #003366;  /* Dark blue for headings */
        }
        .custom-title {
            font-size: 30px;
            color: #003366;
        }
        .fun-icons {
            font-size: 100px;
            color: #4CAF50;
            margin-top: 20px;
            text-align: center;
        }
        .footer {
            text-align: center;
            font-size: 12px;
            color: #003366;
            margin-top: 20px;
        }
        /* Sidebar */
        .sidebar .sidebar-content {
            background-color: #B2C4D8;  /* Light blue sidebar */
        }
        .sidebar .sidebar-content a {
            text-decoration: none;
            color: #003366;
            display: block;
            padding: 10px;
            margin: 5px;
            border-radius: 5px;
            background-color: #FFFFFF;
            font-size: 16px;
        }
        .sidebar .sidebar-content a:hover {
            background-color: #4CAF50;
            color: white;
        }
    </style>
""", unsafe_allow_html=True)

# Sidebar Configuration
st.sidebar.title("📂 Options")
uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

# Sidebar Sections for navigation
sections = {
    "Average per Learner": "overview",
    "Pie Charts Analysis": "charts",
    "Insights and Recommendations": "recommendations",
    "Download Full Report": "download"
}

# Sidebar navigation using links (Clickable Bullet Points)
for section_name, section_id in sections.items():
    st.sidebar.markdown(f'<a href="#{section_id}" class="custom-button">{section_name}</a>', unsafe_allow_html=True)

# Upload file and process
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Check for necessary columns and remove unwanted ones
    if 'NO' in df.columns:
        df = df.drop(columns=['NO'])

    # Detect name column and numeric columns
    name_col = None
    for col in df.columns:
        if "name" in col.lower():
            name_col = col
            break

    numeric_cols = df.select_dtypes(include=['number']).columns
    question_cols = [col for col in numeric_cols if 'q' in col.lower() or 'question' in col.lower()]

    # Calculate Average once for all sections
    if name_col and question_cols:
        df['Average'] = df[question_cols].mean(axis=1)

    # Section: Results Overview (Average per Learner)
    if st.markdown(f'<a name="overview"></a>', unsafe_allow_html=True):
        st.markdown("<h1 class='custom-title' style='text-align: center;'>Saul Damon High School</h1>", unsafe_allow_html=True)
        st.subheader("Results DASHBOARD")
        st.write("Upload your class mark sheet to generate insightful analysis.")

        # Overview: Average per Learner
        st.subheader("📊 Average Mark per Learner")
        if name_col:
            avg_fig, ax1 = plt.subplots(figsize=(8, 5))  # Smaller chart size
            sns.barplot(x=df[name_col], y='Average', data=df.sort_values('Average'), ax=ax1, palette="Blues_d")
            ax1.set_title("Average Marks per Learner", color='#003366')
            ax1.set_xlabel("Learner")
            ax1.set_ylabel("Average Mark")
            ax1.set_xticklabels(ax1.get_xticklabels(), rotation=45, ha='right')
            st.pyplot(avg_fig)  # Display the chart on the page

    # Section: Pie Charts
    if st.markdown(f'<a name="charts"></a>', unsafe_allow_html=True):
        st.subheader("📊 Pie Chart Analysis per Question")
        if question_cols:
            for question in question_cols:
                pie_fig, ax2 = plt.subplots(figsize=(5, 5))  # Smaller pie chart
                question_data = df[question].value_counts()
                ax2.pie(question_data, labels=question_data.index, autopct='%1.1f%%', startangle=90, colors=sns.color_palette("Set2", len(question_data)))
                ax2.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
                ax2.set_title(f"Distribution of Marks for {question}", color='#003366')
                st.pyplot(pie_fig)  # Display pie charts for each question

    # Section: Insights and Recommendations
    if st.markdown(f'<a name="recommendations"></a>', unsafe_allow_html=True):
        st.subheader("📝 Insights and Recommendations")
        insights = []
        recommendations = []

        if name_col and question_cols:
            avg_class_mark = df['Average'].mean()
            max_mark = df[question_cols].max().mean()
            avg_percentage = (avg_class_mark / max_mark) * 100
            insights.append(f"The class average is {avg_percentage:.2f}%, indicating overall performance.")
            low_performers = df[df['Average'] < avg_class_mark * 0.8][name_col].tolist()
            if low_performers:
                insights.append(f"Learners below 80% of class average: {', '.join(low_performers)}.")
                recommendations.append("Consider targeted interventions or extra support for learners struggling below the 80% threshold.")

            question_difficulty = df[question_cols].mean()[df[question_cols].mean() < df[question_cols].mean().mean() * 0.9]
            if not question_difficulty.empty:
                insights.append(f"Challenging questions (below 90% of mean): {', '.join(question_difficulty.index)}.")
                recommendations.append("Review teaching methods or materials for these topics to improve understanding.")

            if df['Average'].std() > 10:
                insights.append("High variability in learner performance (std > 10), suggesting inconsistent understanding.")
                recommendations.append("Implement peer tutoring or group work to balance performance across learners.")

        st.write("**Insights:**")
        for insight in insights:
            st.write(f"- {insight}")
        st.write("**Recommendations:**")
        for rec in recommendations:
            st.write(f"- {rec}")

    # Section: Download Full Report
    if st.markdown(f'<a name="download"></a>', unsafe_allow_html=True):
        st.subheader("⬇️ Download Full Report")

        # Function to generate PDF report
        def generate_pdf():
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()

            # Title
            pdf.set_font("Arial", "B", 16)
            pdf.cell(200, 10, txt="Saul Damon High School", ln=True, align='C')
            pdf.cell(200, 10, txt="Results DASHBOARD", ln=True, align='C')
            pdf.ln(10)

            # Insights
            pdf.set_font("Arial", "B", 12)
            pdf.cell(200, 10, txt="Insights", ln=True)
            pdf.set_font("Arial", size=12)
            for insight in insights:
                pdf.multi_cell(0, 10, txt=f"- {insight}")
            pdf.ln(5)

            # Recommendations
            pdf.set_font("Arial", "B", 12)
            pdf.cell(200, 10, txt="Recommendations", ln=True)
            pdf.set_font("Arial", size=12)
            for rec in recommendations:
                pdf.multi_cell(0, 10, txt=f"- {rec}")
            pdf.ln(5)

            # Save charts as images and include them in the PDF
            def save_chart_to_image(fig, path):
                fig.savefig(path, format='png')
                plt.close(fig)

            # Save the average chart
            avg_image_path = "avg_chart.png"
            avg_fig, ax1 = plt.subplots(figsize=(8, 5))
            sns.barplot(x=df[name_col], y='Average', data=df.sort_values('Average'), ax=ax1, palette="Blues_d")
            ax1.set_title("Average Marks per Learner", color='#003366')
            ax1.set_xlabel("Learner")
            ax1.set_ylabel("Average Mark")
            ax1.set_xticklabels(ax1.get_xticklabels(), rotation=45, ha='right')
            save_chart_to_image(avg_fig, avg_image_path)
            pdf.image(avg_image_path, x=None, y=None, w=150)

            # Save Pie charts for each question
            for question in question_cols:
                pie_image_path = f"pie_chart_{question}.png"
                pie_fig, ax2 = plt.subplots(figsize=(5, 5))
                question_data = df[question].value_counts()
                ax2.pie(question_data, labels=question_data.index, autopct='%1.1f%%', startangle=90, colors=sns.color_palette("Set2", len(question_data)))
                ax2.axis('equal')
                ax2.set_title(f"Distribution of Marks for {question}", color='#003366')
                save_chart_to_image(pie_fig, pie_image_path)
                pdf.image(pie_image_path, x=None, y=None, w=150)

            return pdf.output(dest='S').encode('latin1')

        if st.button("Generate and Download PDF"):
            pdf_data = generate_pdf()
            st.download_button(label="Download PDF Report", data=pdf_data, file_name="performance_report.pdf", mime="application/pdf")

    # Footer
    st.markdown("<div class='footer'>Designed by Mr AR Visagie</div>", unsafe_allow_html=True)

