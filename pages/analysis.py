import streamlit as st
import pandas as pd
import plotly.express as px
import seaborn as sns
import matplotlib.pyplot as plt

# Inject favicon
st.markdown(
    '<link rel="icon" type="image/x-icon" href="./favicon.ico">',
    unsafe_allow_html=True
)

# Custom Styling
st.markdown("""
    <style>
        .stApp { background-color: #F4F6F9; color: #333333; font-family: Arial, sans-serif; }
        .custom-title { font-size: 30px; color: #003366; }
        h2, h3 { color: #003366; }
        .stButton>button { background-color: #003366; color: white; }
    </style>
""", unsafe_allow_html=True)

st.title("Learner Performance Analysis")

# Check session state
if not st.session_state.get('file_processed', False):
    st.error("Please upload an Excel file on the Home page first.")
    st.stop()

# Retrieve data
df = st.session_state['df']
name_col = st.session_state['name_col']

# Re-filter question_cols to ensure only active questions are included
all_question_cols = [col for col in df.columns if df[col].dtype in ['int64', 'float64'] and col not in ['Total', 'Percentage']]
question_cols = [col for col in all_question_cols if df[col].sum() > 0]  # Only questions with non-zero sums

# Ensure Total and Percentage are calculated based on active questions only
if 'Total' not in df.columns or 'Percentage' not in df.columns:
    uploaded_file = st.session_state['uploaded_file']
    df_raw = pd.read_excel(uploaded_file, header=None)
    total_row = next((i for i, row in df_raw.iterrows() if any("totaal:" in str(cell).lower() for cell in row)), None)
    if total_row is not None:
        total_value = df_raw.iloc[total_row].apply(str).str.extract(r'(\d+)').dropna().iloc[0, 0]
        max_possible = int(total_value)
    else:
        max_possible = 50
    df['Total'] = df[question_cols].sum(axis=1)
    df['Percentage'] = (df['Total'] / max_possible * 100)
    st.session_state['df'] = df  # Update session state
    st.session_state['question_cols'] = question_cols  # Update stored question_cols

# Create tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Learner Dashboard",
    "Question Analysis",
    "Progress Tracker",
    "Comparative Analysis",
    "Interactive Insights"
])

# Tab 1: Learner Dashboard
with tab1:
    st.subheader("Individual Learner Dashboard")
    st.markdown("This section shows a selected learner’s strengths and weaknesses compared to the class average.")
    learner = st.selectbox("Select Learner", df[name_col])
    learner_data = df[df[name_col] == learner].iloc[0]

    radar_fig = px.line_polar(
        r=learner_data[question_cols].values,
        theta=question_cols,
        line_close=True,
        title=f"Performance Profile for {learner}"
    )
    radar_fig.update_traces(fill='toself', line_color='#003366')
    st.plotly_chart(radar_fig)
    st.markdown("**Radar Chart Explanation:** This shows the learner’s marks per active question.")

    avg_class = df[question_cols].mean()
    comparison_df = pd.DataFrame({
        "Question": question_cols,
        "Learner": learner_data[question_cols].values,
        "Class Average": avg_class.values
    }).melt(id_vars=["Question"], var_name="Category", value_name="Marks")
    bar_fig = px.bar(
        comparison_df,
        x="Question",
        y="Marks",
        color="Category",
        barmode='group',
        title=f"{learner} vs. Class Average",
        color_discrete_map={"Learner": "#003366", "Class Average": "#4CAF50"}
    )
    st.plotly_chart(bar_fig)
    st.markdown("**Bar Chart Explanation:** Compares the learner’s mark to the class average for each active question.")
    weak_q = learner_data[question_cols].idxmin()
    st.markdown(f"**Focus Area:** {weak_q} (Score: {learner_data[weak_q]:.1f})")

# Tab 2: Question Analysis
with tab2:
    st.subheader("Question Performance")
    st.markdown("This section shows performance details for each active question.")
    question = st.selectbox("Select Question", question_cols)
    box_fig = px.box(df, y=question, title=f"Distribution for {question}", color_discrete_sequence=['#003366'])
    st.plotly_chart(box_fig)
    st.markdown("**Box Plot Explanation:** Shows spread of marks for the selected question.")

    st.markdown("### Top 5 Weakest Questions")
    avg_scores = df[question_cols].mean().sort_values().head(5)
    weak_fig = px.bar(
        avg_scores,
        orientation='v',
        labels={'value': 'Average Score', 'index': 'Question'},
        title="Questions with Lowest Class Averages",
        color_discrete_sequence=['#D32F2F']
    )
    st.plotly_chart(weak_fig)
    st.markdown("**Chart Explanation:** The 5 questions with the lowest average scores.")
    stats = df[question].describe()
    st.markdown(f"**Stats for {question}:**")
    st.write(stats)

# Tab 3: Progress Tracker
with tab3:
    st.subheader("Progress Over Time")
    st.markdown("Shows how average marks change over time. Requires a 'Test Date' column.")
    if 'Test Date' in df.columns:
        df['Test Date'] = pd.to_datetime(df['Test Date'])
        progress_df = df.groupby('Test Date')[question_cols].mean().reset_index()
        progress_fig = px.line(progress_df, x='Test Date', y=question_cols, title="Average Marks Over Time")
        st.plotly_chart(progress_fig)
        st.markdown("**Line Chart Explanation:** Tracks the change in class average for each active question over time.")
    else:
        st.warning("Your current Excel file lacks a 'Test Date' column. Add it to track progress.")

# Tab 4: Comparative Analysis
with tab4:
    st.subheader("Group Comparison")
    st.markdown("Upload another file to compare your group’s performance.")
    additional_file = st.file_uploader("Upload comparison Excel file", type=["xlsx"])
    if additional_file:
        df2_raw = pd.read_excel(additional_file, header=None)
        header_row2 = next((i for i, row in df2_raw.iterrows() if any("name of learner" in str(cell).lower() for cell in row)), None)
        total_row2 = next((i for i, row in df2_raw.iterrows() if any("totaal:" in str(cell).lower() for cell in row)), None)
        if header_row2 is not None:
            df2 = pd.read_excel(additional_file, skiprows=header_row2)
            df2.columns = df2.columns.str.strip()
            name_col2 = next((col for col in df2.columns if "name of learner" in col.lower()), None)
            if name_col2:
                name_col_idx2 = df2.columns.get_loc(name_col2)
                all_q2 = [col for col in df2.columns[name_col_idx2 + 1:] if df2[col].dtype in ['int64', 'float64']]
                df2[all_q2] = df2[all_q2].fillna(0)
                q2 = [col for col in all_q2 if df2[col].sum() > 0]  # Filter active questions
                df2['Total'] = df2[q2].sum(axis=1)
                if total_row2 is not None:
                    total_value2 = df2_raw.iloc[total_row2].apply(str).str.extract(r'(\d+)').dropna().iloc[0, 0]
                    max_possible2 = int(total_value2)
                else:
                    max_possible2 = 50
                df2['Percentage'] = (df2['Total'] / max_possible2 * 100)
                common_q = [q for q in question_cols if q in q2]
                if common_q:
                    compare_fig = px.bar(
                        pd.DataFrame({"Original Group": df[common_q].mean(), "New Group": df2[common_q].mean()}).reset_index().rename(columns={"index": "Question"}),
                        x="Question",
                        y=["Original Group", "New Group"],
                        barmode='group',
                        title="Average Marks Comparison",
                        color_discrete_map={"Original Group": "#003366", "New Group": "#4CAF50"}
                    )
                    st.plotly_chart(compare_fig)
                    st.markdown("**Bar Chart Explanation:** Shows how two groups performed on the same active questions.")
                else:
                    st.error("No matching active question columns between the two files.")
            else:
                st.error("Could not find 'NAME OF LEARNER' in the second file.")
        else:
            st.error("Could not locate 'NAME OF LEARNER' row in the second file.")
    else:
        st.markdown("**Current Group Averages**")
        st.bar_chart(df[question_cols].mean())

# Tab 5: Interactive Insights
with tab5:
    st.subheader("Build Your Own Insights")
    st.markdown("Select active questions and a chart type to explore the data.")
    selected_questions = st.multiselect("Select Questions", question_cols, default=question_cols[:2])
    chart_type = st.selectbox("Chart Type", ["Bar", "Scatter", "Histogram"], index=0)
    if selected_questions:
        if chart_type == "Bar":
            bar_fig = px.bar(df, x=name_col, y=selected_questions, barmode='group', title="Custom Bar Chart")
            bar_fig.update_layout(xaxis={'tickangle': 90})
            st.plotly_chart(bar_fig)
            st.markdown("**Bar Chart Explanation:** Each learner’s performance for selected active questions.")
        elif chart_type == "Scatter" and len(selected_questions) >= 2:
            scatter_fig = px.scatter(df, x=selected_questions[0], y=selected_questions[1], hover_data=[name_col], title=f"{selected_questions[0]} vs {selected_questions[1]}")
            st.plotly_chart(scatter_fig)
            st.markdown("**Scatter Plot Explanation:** Compares two selected active questions.")
        elif chart_type == "Histogram":
            hist_fig = px.histogram(df, x=selected_questions[0], title=f"Histogram of {selected_questions[0]}")
            st.plotly_chart(hist_fig)
            st.markdown("**Histogram Explanation:** Shows distribution of scores for the selected active question.")
        else:
            st.warning("Select at least 2 questions for a Scatter plot.")

st.markdown('<div class="footer">Designed by Mr AR Visagie</div>', unsafe_allow_html=True)