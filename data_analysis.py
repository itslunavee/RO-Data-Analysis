import pandas as pd
from collections import Counter
import matplotlib.pyplot as plt

ISSUE_CATEGORIES = {
    "FASA": ["fasa", "financial aid", "part-time"],
    "Login": ["login", "password", "authentication"],
    "International": ["international", "visa", "study permit"],
    "Technical": ["system", "error", "bug", "technical"],
    "Registration": ["register", "course", "enrollment"]
}

def categorize_text(text):
    text = str(text).lower()
    for category, keywords in ISSUE_CATEGORIES.items():
        if any(keyword in text for keyword in keywords):
            return category
    return "Other"

def analyze_frequencies(df):
    #top recurring phrases
    from collections import Counter
    words = ' '.join(df['Description'].dropna()).lower().split()
    bigrams = Counter(zip(words, words[1:]))
    
    #most common key words
    word_counts = Counter(words)
    
    return {
        'top_bigrams': bigrams.most_common(20),
        'top_words': word_counts.most_common(30)
    }

def load_data():
    df = pd.read_excel("../Data/Output/cleaned_output.xlsx")
    
    if 'Date/Time Closed' in df.columns:
        df['Date'] = pd.to_datetime(
            df['Date/Time Closed'].str.replace(r'([ap])\.m\.', r'\1m', regex=True),
            format='%Y-%m-%d, %I:%M %p', 
            errors='coerce'
        )
    return df

def data_analysis(df):
    #category distribution
    df['Category'] = df['Description'].apply(categorize_text)
    category_counts = df['Category'].value_counts()
    
    #time trends
    time_trends = None
    if 'Date/Time Closed' in df.columns:
        time_trends = df.groupby([df['Date'].dt.to_period('M'), 'Category']).size()
    
    #common phrases
    freq_analysis = analyze_frequencies(df)
    
    return {
        'category_counts': category_counts,
        'time_trends': time_trends,
        'common_phrases': freq_analysis
    }

def visualize_results(results):
    #pie chart
    results['category_counts'].plot.pie(autopct='%1.1f%%')
    plt.title("Issue Category Distribution")
    plt.show()
    
    #time trends chart
    if 'time_trends' in results and not results['time_trends'].empty:
        results['time_trends'].unstack().plot(kind='bar', stacked=True)
        plt.title("Issues by Month")
        plt.tight_layout() 
        plt.show()

def save_analysis(results, output_file="../Analysis/analysis_report.xlsx"):
    with pd.ExcelWriter(output_file) as writer:
        #save category counts
        results['category_counts'].to_excel(writer, sheet_name='Categories')
        
        #save time trends
        if 'time_trends' in results and not results['time_trends'].empty:
            results['time_trends'].unstack().to_excel(writer, sheet_name='Time Trends')
        
        #saves common phrases
        pd.DataFrame(results['common_phrases']['top_bigrams'], 
                    columns=['Phrase', 'Count']).to_excel(writer, sheet_name='Common Phrases')

if __name__ == "__main__":
    print("🔍 Loading data...")
    df = load_data()
    
    print("📊 Analyzing...")
    results = data_analysis(df)
    
    print("📈 Visualizing...")
    visualize_results(results)
    
    print("💾 Saving results...")
    save_analysis(results)
    
    print("✅ Analysis complete!")