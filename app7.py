import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
import random

# Function to load data
def load_data(file_path):
    try:
        df = pd.read_excel(file_path)
        st.write("Data loaded successfully.")
        return df
    except FileNotFoundError:
        st.error(f"The file was not found at {file_path}")
        return None
    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None

# Function to handle user queries
def handle_query(df, query):
    query = query.lower()

    # Define keyword groups
    top_keywords = ['top', 'best', 'highest', 'top 5']
    bottom_keywords = ['bottom', 'worst', 'lowest', 'bottom 5']
    company_keywords = ['company', 'companies']
    revenue_keywords = ['revenue', 'sales', 'total revenue', 'maximum revenue', 'total number of sales']
    profit_keywords = ['profit', 'income', 'earnings', 'total profit', 'minimum profit']
    count_keywords = ['how many', 'number of', 'count']
    random_keywords = ['random', 'any']

    # Handling queries for the top 5 companies by revenue or profit
    if any(keyword in query for keyword in top_keywords) and any(keyword in query for keyword in company_keywords):
        if any(keyword in query for keyword in revenue_keywords):
            top_5_revenue = df.nlargest(5, 'Revenue')[['Company', 'Revenue']]
            st.write("Top 5 companies by revenue:")
            st.write(top_5_revenue)
        elif any(keyword in query for keyword in profit_keywords):
            top_5_profit = df.nlargest(5, 'Profit')[['Company', 'Profit']]
            st.write("Top 5 companies by profit:")
            st.write(top_5_profit)
        else:
            top_5_revenue = df.nlargest(5, 'Revenue')[['Company', 'Revenue']]
            st.write("Top 5 companies by revenue:")
            st.write(top_5_revenue)
        return True

    # Handling queries for the bottom 5 companies by revenue or profit
    elif any(keyword in query for keyword in bottom_keywords) and any(keyword in query for keyword in company_keywords):
        if any(keyword in query for keyword in revenue_keywords):
            bottom_5_revenue = df.nsmallest(5, 'Revenue')[['Company', 'Revenue']]
            st.write("Bottom 5 companies by revenue:")
            st.write(bottom_5_revenue)
        elif any(keyword in query for keyword in profit_keywords):
            bottom_5_profit = df.nsmallest(5, 'Profit')[['Company', 'Profit']]
            st.write("Bottom 5 companies by profit:")
            st.write(bottom_5_profit)
        else:
            bottom_5_revenue = df.nsmallest(5, 'Revenue')[['Company', 'Revenue']]
            st.write("Bottom 5 companies by revenue:")
            st.write(bottom_5_revenue)
        return True

    # Handling queries for total number of companies
    elif any(keyword in query for keyword in count_keywords) and any(keyword in query for keyword in company_keywords):
        num_companies = df['Company'].nunique()
        st.write(f"Total number of companies: {num_companies}")
        return True

    # Handling queries for highest profit
    elif 'highest' in query and 'profit' in query:
        if 'company' in query:
            highest_profit = df.loc[df['Profit'].idxmax()][['Company', 'Profit']]
            st.write("Company with the highest profit:")
            st.write(highest_profit)
        elif 'companies' in query:
            top_5_profit = df.nlargest(5, 'Profit')[['Company', 'Profit']]
            st.write("Top 5 companies by profit:")
            st.write(top_5_profit)
        return True

    # Handling queries for maximum revenue
    elif 'maximum' in query and 'revenue' in query:
        if 'company' in query:
            max_revenue = df.loc[df['Revenue'].idxmax()][['Company', 'Revenue']]
            st.write("Company with the maximum revenue:")
            st.write(max_revenue)
        elif 'companies' in query:
            top_5_revenue = df.nlargest(5, 'Revenue')[['Company', 'Revenue']]
            st.write("Top 5 companies by revenue:")
            st.write(top_5_revenue)
        return True

    # Handling queries for minimum profit
    elif 'minimum' in query and 'profit' in query:
        if 'company' in query:
            min_profit = df.loc[df['Profit'].idxmin()][['Company', 'Profit']]
            st.write("Company with the minimum profit:")
            st.write(min_profit)
        elif 'companies' in query:
            bottom_5_profit = df.nsmallest(5, 'Profit')[['Company', 'Profit']]
            st.write("Bottom 5 companies by profit:")
            st.write(bottom_5_profit)
        return True

    # Handling queries for total revenue
    elif 'total' in query and 'revenue' in query:
        total_revenue = df['Revenue'].sum()
        st.write(f"Total revenue of all companies: {total_revenue}")
        return True

    # Handling queries for total profit
    elif 'total' in query and 'profit' in query:
        total_profit = df['Profit'].sum()
        st.write(f"Total profit of all companies: {total_profit}")
        return True

    # Handling queries for random companies
    elif any(keyword in query for keyword in random_keywords):
        random_companies = df.sample(n=5)[['Company', 'Revenue', 'Profit']]
        st.write("Here are 5 random companies:")
        st.write(random_companies)
        return True

    # If the query contains a specific company name
    else:
        company_name = query.strip().title()  # Capitalize each word for consistent matching
        if company_name in df['Company'].values:
            display_company_info(df, company_name)
            plot_charts(df, company_name)
            return True

    # If no matching query is found
    return False

# Function to display company information
def display_company_info(df, company_name):
    company_data = df[df['Company'] == company_name]
    if company_data.empty:
        st.write(f"No data found for company: {company_name}")
        return None
    else:
        st.write(f"Data for company: {company_name}")
        st.write(company_data)
        return company_data

# Function to plot charts
def plot_charts(df, company_name):
    company_data = df[df['Company'] == company_name]

    if company_data.empty:
        st.write("No data to display for the specified company.")
        return

    # Remove duplicate entries for charting
    unique_df = df.drop_duplicates(subset=['Company'])

    # Bar Chart: User Input vs. Top 5 Companies by Revenue
    st.subheader('Revenue Model')
    top_5 = unique_df.nlargest(5, 'Revenue')
    top_5_and_user = pd.concat([top_5, company_data]).drop_duplicates(subset=['Company'])

    plt.figure(figsize=(10, 6))
    bar_plot = sns.barplot(x='Company', y='Revenue', data=top_5_and_user, palette='viridis')
    plt.title('Revenue Model', color='white', fontsize=20)
    plt.xticks(rotation=30, ha='center', color='red')  # Rotate labels for better readability
    plt.tight_layout()

    for p in bar_plot.patches:
        bar_plot.annotate(format(p.get_height(), '.0f'),
                          (p.get_x() + p.get_width() / 2., p.get_height()),
                          ha='center', va='center', xytext=(0, 10),
                          textcoords='offset points', color='red', fontsize=12, fontweight='bold')

    st.pyplot(plt)

    # Pie Chart: User Input vs. Top 5 Companies by Profit
    st.subheader('Market Share')
    top_5_profit = unique_df.nlargest(5, 'Profit')
    top_5_profit_and_user = pd.concat([top_5_profit, company_data]).drop_duplicates(subset=['Company'])

    plt.figure(figsize=(8, 8))
    plt.pie(top_5_profit_and_user['Profit'], labels=top_5_profit_and_user['Company'], autopct='%1.1f%%', colors=sns.color_palette("dark:#5A9_r", len(top_5_profit_and_user)))
    plt.title('Market Share', color='white', fontsize=20)
    plt.tight_layout()
    st.pyplot(plt)

    # Histogram: Distribution of Profit
    st.subheader('Profit Distribution')
    plt.figure(figsize=(10, 6))
    bins = 20
    counts, bin_edges, patches = plt.hist(df['Profit'], bins=bins, color='skyblue', edgecolor='black', alpha=0.7, label='Profit Distribution')
    plt.title('Profit Distribution', color='white', fontsize=20)
    plt.xlabel('Profit', color='white')
    plt.ylabel('Frequency', color='white')

    # Add data labels to the histogram
    for count, bin_edge, patch in zip(counts, bin_edges, patches):
        height = patch.get_height()
        plt.text(bin_edge + (patch.get_width() / 2), height, f'{int(height)}',
                 ha='center', va='bottom', color='red', fontsize=12, fontweight='bold')

    plt.tight_layout()
    st.pyplot(plt)

# Streamlit app
def main():
    st.title('Excel Data Insights Chatbot')

    # File path for the Excel file
    file_path = "C:\\Users\\Joydip\\Downloads\\Data212.xlsx"

    # Load data from the specified file path
    df = load_data(file_path)
    if df is None:
        st.stop()  # Stop the app if no data is loaded

    # Show data if loaded successfully
    if df is not None:
        st.write("Data loaded successfully. Ask your question below.")

    # Text input for the query
    query = st.text_input("Ask your query here:")

    # Process the query when the user inputs something
    if query:
        result = handle_query(df, query)
        if not result:
            st.write("Sorry, I couldn't understand your query.")
if __name__ == '__main__':
    main()
