import pandas as pd
from textblob import TextBlob
from langdetect import detect
from googletrans import Translator
import spacy
import warnings
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import textwrap
import time
import re
from collections import defaultdict
from datetime import datetime

#pip install pandas textblob googletrans langdetect spacy matplotlib seaborn openpyxl xlrd python-dotenv tqdm && python -m spacy download en_core_web_md
#(Windows10/11)pip install pywin32==306        # Required for some Excel operations

# Suppress warnings
warnings.filterwarnings("ignore", category=UserWarning)

# Load spaCy model
nlp = spacy.load("en_core_web_md")
translator = Translator()

# Modern color scheme
COLORS = {
    "background": "#f5f5f5",
    "primary": "#6200ee",
    "primary_light": "#9e47ff",
    "secondary": "#03dac6",
    "text": "#333333",
    "card": "#ffffff",
    "border": "#e0e0e0"
}

# Define categories
categories = [
    "Business Skills",
    "Technical Skills",
    "Leadership",
    "Quality Risk and Management"
]

# Refined category keywords based on real feedback
category_keywords = {
    "Business Skills": ["business", "market", "sales", "strategy", "client", "revenue", "opportunity", "crm", "gtm"],
    "Technical Skills": ["technical", "software", "engineering", "tools", "systems", "code", "development", "data", "technology"],
    "Leadership": ["lead", "team", "mentor", "manage", "vision", "delegate", "inspire", "guide", "encadrer", "équipe"],
    "Quality Risk and Management": ["quality", "risk", "compliance", "process", "control", "governance", "standards"]
}

# Feedback response scores
feedback_response_scores = {
    "Significantly exceeded expectations": 5,
    "Exceeded expectations": 4,
    "Met expectations": 3,
    "Partially met expectations": 2,
    "Did not meet expectations": 1,
    "Non observe": 0,
    "Not observed": 0,
    "Business skills": 3,
    "Technical skills": 3,
    "Leadership skills": 3
}

# Translation cache
translation_cache = {}

def normalize_text(text):
    if not isinstance(text, str):
        return ""
    return re.sub(r'[^\w\s]', '', text.lower())

def translate_text(text):
    """Translate French text to English with safety checks."""
    if not isinstance(text, str) or text.strip() == "":
        return text
    if text in translation_cache:
        return translation_cache[text]
    
    try:
        lang = detect(text.strip())
        if lang == 'fr':
            result = translator.translate(text.strip(), src='fr', dest='en').text
            translation_cache[text] = result
            return result
        else:
            return text.strip()
    except Exception as e:
        print(f"Translation error: {e}")
        return text.strip()

def classify_text(text):
    """Hybrid scoring using both spaCy similarity and keyword frequency."""
    if not text.strip():
        return {cat: 0 for cat in categories}
    
    # Keyword-based score
    words = text.lower().split()
    word_count = len(words)

    keyword_scores = {}
    for category, keywords in category_keywords.items():
        count = sum(1 for word in words if word in keywords)
        score = count / word_count if word_count > 0 else 0
        keyword_scores[category] = score
    
    # SpaCy similarity score
    doc = nlp(text)
    spacy_scores = {category: doc.similarity(nlp(category)) for category in categories}

    # Combined hybrid score
    return {
        category: 0.5 * spacy_scores[category] + 0.5 * keyword_scores[category]
        for category in categories
    }

def count_and_score_feedback_responses(responses):
    """Count how many times each feedback response appears."""
    response_counts = {response: 0 for response in feedback_response_scores}
    total_score = 0
    total_responses = 0

    for response in responses:
        if not isinstance(response, str) or not response.strip():
            continue
        for key in response_counts:
            if key.lower() in response.lower():
                response_counts[key] += 1
                total_score += feedback_response_scores[key]
                total_responses += 1
                break

    avg_score = total_score / total_responses if total_responses > 0 else 0
    return response_counts, avg_score

def process_feedback(feedback_id, comments, responses, questions, ranks):
    """Aggregate and analyze feedback data with rank-based weighting."""
    # Count and score feedback responses
    response_counts, avg_score = count_and_score_feedback_responses(responses)
    
    # Prepare response analysis text
    response_text = "Feedback Response Analysis:\n"
    for response, count in response_counts.items():
        response_text += f"{response}: {count} (Score: {feedback_response_scores[response]})\n"
    response_text += f"\nAverage Feedback Score: {avg_score:.2f}/5.00\n"
    response_text += "-"*40 + "\n"

    total_sentiment = 0
    category_scores = {category: 0 for category in categories}
    question_analysis = {}
    rank_analysis = {}
    weighted_count = 0

    # Process comments and responses
    for comment, response, question, rank in zip(comments, responses, questions, ranks):
        full_text = ""
        if isinstance(comment, str) and comment.strip():
            full_text += translate_text(comment)
        if isinstance(response, str) and response.strip():
            full_text += " " + translate_text(response)

        if not full_text.strip():
            continue

        sentiment = TextBlob(full_text).sentiment.polarity
        scores = classify_text(full_text)

        for cat in scores:
            category_scores[cat] += scores[cat]
        total_sentiment += sentiment
        weighted_count += 1

        # Analyze by question
        if question not in question_analysis:
            question_analysis[question] = {
                "texts": [],
                "sentiment": 0,
                "category_scores": {category: 0 for category in categories},
                "weighted_count": 0
            }
        question_analysis[question]["texts"].append(full_text)
        question_analysis[question]["sentiment"] += sentiment
        for cat in scores:
            question_analysis[question]["category_scores"][cat] += scores[cat]
        question_analysis[question]["weighted_count"] += 1

        # Analyze by rank
        if rank not in rank_analysis:
            rank_analysis[rank] = {
                "texts": [],
                "sentiment": 0,
                "category_scores": {category: 0 for category in categories},
                "weighted_count": 0
            }
        rank_analysis[rank]["texts"].append(full_text)
        rank_analysis[rank]["sentiment"] += sentiment
        for cat in scores:
            rank_analysis[rank]["category_scores"][cat] += scores[cat]
        rank_analysis[rank]["weighted_count"] += 1

    if weighted_count == 0:
        return "No valid feedback found.", None, None, None, None

    # Average scores
    avg_sentiment = total_sentiment / weighted_count
    avg_category_scores = {cat: score / weighted_count for cat, score in category_scores.items()}
    highest_category = max(avg_category_scores, key=avg_category_scores.get)

    # Build main output
    main_text = response_text
    main_text += "Main Analysis:\n"
    main_text += f"Average Sentiment: {avg_sentiment:.2f}\n"
    main_text += "Category Similarity Scores:\n"
    for category, score in avg_category_scores.items():
        main_text += f"  {category}: {score:.2f}\n"
    main_text += f"Most Relevant Category: {highest_category}\n"
    main_text += "-"*40 + "\n"

    # Detailed breakdown
    detailed_text = f"Detailed Breakdown for ID: {feedback_id}\n"
    for comment, response, rank in zip(comments, responses, ranks):
        if comment.strip() or response.strip():
            detailed_text += f"[Rank: {rank}]\n"
            if comment.strip():
                detailed_text += f"Comment: {comment}\n"
            if response.strip():
                detailed_text += f"Response: {response}\n"
            detailed_text += "---\n"

    return main_text, detailed_text, avg_category_scores, question_analysis, rank_analysis

def show_chart_window(similarity_scores):
    chart_window = tk.Toplevel()
    chart_window.title("Feedback Analysis Charts")
    chart_window.geometry("1000x600")
    chart_window.configure(bg=COLORS['background'])

    chart_frame = tk.Frame(chart_window, bg=COLORS['card'], bd=2, relief='groove')
    chart_frame.pack(fill='both', expand=True, padx=20, pady=20)

    if similarity_scores:
        try:
            plt.style.use('seaborn')
        except:
            plt.style.use('default')

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
        categories_list = ["\n".join(textwrap.wrap(cat, width=10)) for cat in categories]
        scores = [similarity_scores.get(category, 0) for category in categories]

        bars = ax1.bar(categories_list, scores, color=COLORS['primary_light'])
        ax1.set_facecolor(COLORS['card'])
        fig.patch.set_facecolor(COLORS['card'])
        ax1.spines['top'].set_visible(False)
        ax1.spines['right'].set_visible(False)
        ax1.set_xlabel("Categories", fontsize=10)
        ax1.set_ylabel("Similarity Score", fontsize=10)
        ax1.set_title("Category Similarity Scores", fontsize=12, pad=20)
        ax1.set_ylim(0, 1)

        wedges, texts, autotexts = ax2.pie(
            scores, labels=categories_list, autopct='%1.1f%%',
            startangle=90, colors=[COLORS['primary'], COLORS['secondary'], '#ffab91', '#ce93d8'],
            textprops={'fontsize': 8}
        )
        ax2.set_title("Score Distribution", fontsize=12, pad=20)

        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
    else:
        tk.Label(chart_frame, text="No data available to display chart.",
                 bg=COLORS['card'], fg=COLORS['text'], font=('Segoe UI', 10)).pack(pady=20)

def show_improvement_chart(file_paths, feedback_id):
    """Show a chart comparing skill improvement across multiple files."""
    if len(file_paths) < 2:
        messagebox.showwarning("Improvement Analysis", "Please select at least 2 files to compare improvement.")
        return

    try:
        # Process each file and store results
        all_results = []
        
        for file_path in file_paths:
            df = pd.read_excel(file_path)
            
            # Try to extract date from filename or use file modification time
            try:
                # Try to get date from filename (format: YYYY-MM-DD or similar)
                date_str = re.search(r'(\d{4}-\d{2}-\d{2})', file_path).group(1)
                file_date = datetime.strptime(date_str, "%Y-%m-%d").date()
            except:
                # Fallback to file modification date
                file_date = datetime.fromtimestamp(os.path.getmtime(file_path)).date()
            
            # Identify columns
            comment_col = 'Comments ' if 'Comments ' in df.columns else 'Comments' if 'Comments' in df.columns else None
            response_col = 'Feedback Responses' if 'Feedback Responses' in df.columns else 'Feedback Response' if 'Feedback Response' in df.columns else None
            
            if not comment_col or not response_col:
                continue
                
            df[comment_col] = df[comment_col].fillna("")
            df[response_col] = df[response_col].fillna("")
            df['Feedback Dimensions & Questions'] = df['Feedback Dimensions & Questions'].fillna("")
            df['Rank feedback provider'] = df['Rank feedback provider'].fillna("Unknown")

            # Filter by feedback ID
            result = df[df['Feedback Requester User ID'].astype(str) == str(feedback_id)]
            
            if not result.empty:
                comments = result[comment_col].tolist()
                responses = result[response_col].tolist()
                questions = result['Feedback Dimensions & Questions'].tolist()
                ranks = result['Rank feedback provider'].tolist()
                
                _, _, similarity_scores, _, _ = process_feedback(
                    feedback_id, comments, responses, questions, ranks
                )
                
                if similarity_scores:
                    all_results.append({
                        'date': file_date,
                        'scores': similarity_scores
                    })
        
        # Sort by date
        all_results.sort(key=lambda x: x['date'])
        
        if len(all_results) < 2:
            messagebox.showwarning("Improvement Analysis", "Not enough data points to show improvement.")
            return
            
        # Create improvement chart
        chart_window = tk.Toplevel()
        chart_window.title(f"Skill Improvement Analysis for ID: {feedback_id}")
        chart_window.geometry("1200x700")
        chart_window.configure(bg=COLORS['background'])

        chart_frame = tk.Frame(chart_window, bg=COLORS['card'], bd=2, relief='groove')
        chart_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        try:
            plt.style.use('seaborn')
        except:
            plt.style.use('default')

        fig, ax = plt.subplots(figsize=(12, 6))
        
        # Prepare data for plotting
        dates = [r['date'] for r in all_results]
        date_labels = [d.strftime('%Y-%m-%d') for d in dates]
        
        # Plot each category
        for category in categories:
            scores = [r['scores'].get(category, 0) for r in all_results]
            ax.plot(date_labels, scores, marker='o', label=category, linewidth=2)
        
        ax.set_facecolor(COLORS['card'])
        fig.patch.set_facecolor(COLORS['card'])
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.set_xlabel("Feedback Date", fontsize=10)
        ax.set_ylabel("Skill Score", fontsize=10)
        ax.set_title(f"Skill Improvement Over Time for ID: {feedback_id}", fontsize=12, pad=20)
        ax.set_ylim(0, 1)
        ax.legend(loc='upper left', bbox_to_anchor=(1, 1))
        ax.grid(True, linestyle='--', alpha=0.6)
        
        plt.xticks(rotation=45)
        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        
        # Add improvement percentages
        info_frame = tk.Frame(chart_window, bg=COLORS['background'])
        info_frame.pack(fill='x', padx=20, pady=(0, 20))
        
        # Calculate improvement percentages
        first_scores = all_results[0]['scores']
        last_scores = all_results[-1]['scores']
        
        improvement_text = "Improvement Percentage:\n"
        for category in categories:
            initial = first_scores.get(category, 0.01)  # Avoid division by zero
            final = last_scores.get(category, 0)
            improvement = ((final - initial) / initial) * 100
            improvement_text += f"{category}: {improvement:+.1f}%\n"
        
        tk.Label(info_frame, text=improvement_text, bg=COLORS['background'], 
                fg=COLORS['text'], font=('Segoe UI', 10), justify='left').pack(side='left')
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create improvement chart: {str(e)}")

def analyze_feedback():
    start_time = time.time()

    # Show loading status
    processing_time_label.config(text="Processing feedback... Please wait.")
    root.update_idletasks()

    # Ask user to select one or more files
    file_paths = filedialog.askopenfilenames(
        title="Select Excel File(s)",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file_paths:
        messagebox.showwarning("File Error", "No file(s) selected.")
        processing_time_label.config(text="Ready")
        return

    feedback_id = id_entry.get().strip()
    if not feedback_id:
        messagebox.showwarning("Input Error", "Please provide a Feedback ID.")
        processing_time_label.config(text="Ready")
        return

    if not feedback_id.isdigit():
        messagebox.showwarning("Input Error", "Feedback ID must be a number.")
        processing_time_label.config(text="Ready")
        return

    feedback_id = int(feedback_id)

    try:
        dfs = []
        for file_path in file_paths:
            df = pd.read_excel(file_path)
            dfs.append(df)

        df_combined = pd.concat(dfs, ignore_index=True)

        # Identify columns
        comment_col = 'Comments ' if 'Comments ' in df_combined.columns else 'Comments' if 'Comments' in df_combined.columns else None
        response_col = 'Feedback Responses' if 'Feedback Responses' in df_combined.columns else 'Feedback Response' if 'Feedback Response' in df_combined.columns else None

        if not comment_col or not response_col:
            messagebox.showwarning("Column Error", "Required columns not found in Excel file(s)")
            processing_time_label.config(text="Ready")
            return

        df_combined[comment_col] = df_combined[comment_col].fillna("")
        df_combined[response_col] = df_combined[response_col].fillna("")
        df_combined['Feedback Dimensions & Questions'] = df_combined['Feedback Dimensions & Questions'].fillna("")
        df_combined['Rank feedback provider'] = df_combined['Rank feedback provider'].fillna("Unknown")

        # Filter by feedback ID
        result = df_combined[df_combined['Feedback Requester User ID'].astype(str) == str(feedback_id)]

        if result.empty:
            main_text = f"No feedback found for ID: {feedback_id}"
            detailed_text = ""
            similarity_scores = None
        else:
            comments = result[comment_col].tolist()
            responses = result[response_col].tolist()
            questions = result['Feedback Dimensions & Questions'].tolist()
            ranks = result['Rank feedback provider'].tolist()

            main_text, detailed_text, similarity_scores, _, _ = process_feedback(
                feedback_id, comments, responses, questions, ranks
            )

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
        processing_time_label.config(text="Ready")
        return

    elapsed_time = time.time() - start_time
    main_text += f"\nProcessing Time: {elapsed_time:.2f} sec"
    main_text += f"\nFiles Processed: {len(file_paths)}"

    # Update GUI
    main_score_text_widget.delete(1.0, tk.END)
    detailed_text_widget.delete(1.0, tk.END)
    main_score_text_widget.insert(tk.END, main_text)
    detailed_text_widget.insert(tk.END, detailed_text)

    # Update charts
    for widget in chart_card.winfo_children():
        widget.destroy()

    if similarity_scores:
        chart_button = ModernButton(chart_card, text="View Full-Screen Chart", command=lambda: show_chart_window(similarity_scores))
        chart_button.pack(pady=10)
        
        # Add improvement analysis button if multiple files selected
        if len(file_paths) >= 2:
            improvement_btn = ModernButton(
                chart_card, 
                text="View Skill Improvement Over Time", 
                command=lambda: show_improvement_chart(file_paths, feedback_id)
            )
            improvement_btn.pack(pady=10)
    else:
        tk.Label(chart_card, text="No chart data available.", bg=COLORS['card'], fg=COLORS['text']).pack()

    processing_time_label.config(text=f"Processed {len(file_paths)} file(s) | Processing Time: {elapsed_time:.2f} sec | Ready")

def create_modern_gui():
    global root
    root = tk.Tk()
    root.title("Feedback Analysis Dashboard")
    root.geometry("1400x800")
    root.configure(bg=COLORS['background'])

    global id_entry, main_score_text_widget, detailed_text_widget, chart_card, processing_time_label

    # Header
    header = tk.Frame(root, bg=COLORS['primary'], height=60)
    header.pack(fill='x')
    title_font = ("Segoe UI", 16, "bold")
    tk.Label(header, text="Feedback Analysis Dashboard", bg=COLORS['primary'], fg='white', font=title_font).pack(side='left', padx=20)

    # Input panel
    input_card = tk.Frame(root, bg=COLORS['card'], padx=20, pady=15, highlightbackground=COLORS['border'], highlightthickness=1)
    input_card.pack(fill='x', padx=20, pady=20)

    id_frame = tk.Frame(input_card, bg=COLORS['card'])
    id_frame.pack(side='left', padx=10, fill='x', expand=True)
    tk.Label(id_frame, text="Feedback ID:", bg=COLORS['card'], fg=COLORS['text']).pack(side='left', padx=5)
    id_entry = tk.Entry(id_frame, width=20, font=('Segoe UI', 10), relief='flat', highlightthickness=1)
    id_entry.pack(side='left', padx=5, fill='x', expand=True)

    analyze_btn = ModernButton(input_card, text="Select Files and Analyze", command=analyze_feedback)
    analyze_btn.pack(side='right', padx=10)

    results_card = tk.Frame(root, bg=COLORS['background'])
    results_card.pack(fill='both', expand=True, padx=20, pady=(0, 20))

    summary_card = tk.Frame(results_card, bg=COLORS['card'], padx=10, pady=10, highlightbackground=COLORS['border'], highlightthickness=1)
    summary_card.pack(side='left', fill='both', expand=True, padx=(0, 10))
    tk.Label(summary_card, text="Summary", bg=COLORS['card'], fg=COLORS['primary'], font=("Segoe UI", 12, "bold")).pack(anchor='w')
    main_score_text_widget = tk.Text(summary_card, wrap='word', bg=COLORS['card'], fg=COLORS['text'], font=("Segoe UI", 10))
    main_score_text_widget.pack(fill='both', expand=True)

    detailed_card = tk.Frame(results_card, bg=COLORS['card'], padx=10, pady=10, highlightbackground=COLORS['border'], highlightthickness=1)
    detailed_card.pack(side='right', fill='both', expand=True, padx=(10, 0))
    tk.Label(detailed_card, text="Detailed Feedback", bg=COLORS['card'], fg=COLORS['primary'], font=("Segoe UI", 12, "bold")).pack(anchor='w')
    detailed_text_widget = tk.Text(detailed_card, wrap='word', bg=COLORS['card'], fg=COLORS['text'], font=("Segoe UI", 10))
    detailed_text_widget.pack(fill='both', expand=True)

    chart_card = tk.Frame(root, bg=COLORS['card'], padx=10, pady=10, highlightbackground=COLORS['border'], highlightthickness=1)
    chart_card.pack(fill='x', padx=20, pady=(0, 20))

    status_bar = tk.Frame(root, bg=COLORS['primary'], height=30)
    status_bar.pack(side='bottom', fill='x')
    processing_time_label = tk.Label(status_bar, text="Ready", bg=COLORS['primary'], fg='white', font=("Segoe UI", 9))
    processing_time_label.pack(side='left', padx=10)

    return root

class ModernText(tk.Text):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.config(
            bg=COLORS['card'],
            fg=COLORS['text'],
            insertbackground=COLORS['primary'],
            selectbackground=COLORS['primary_light'],
            relief='flat',
            padx=10,
            pady=10,
            wrap=tk.WORD,
            font=('Segoe UI', 10)
        )

class ModernButton(tk.Button):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.config(
            bg=COLORS['primary'],
            fg='white',
            activebackground=COLORS['primary_light'],
            activeforeground='white',
            relief='flat',
            padx=20,
            pady=8,
            font=('Segoe UI', 10, 'bold'),
            cursor='hand2'
        )
        self.bind("<Enter>", lambda e: self.config(bg=COLORS['primary_light']))
        self.bind("<Leave>", lambda e: self.config(bg=COLORS['primary']))

class ModernEntry(tk.Entry):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.config(
            bg='white',
            fg=COLORS['text'],
            relief='flat',
            highlightthickness=1,
            highlightcolor=COLORS['primary'],
            highlightbackground=COLORS['border'],
            insertbackground=COLORS['primary'],
            font=('Segoe UI', 10),
            selectbackground=COLORS['primary_light']
        )

class ModernLabel(tk.Label):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.config(
            bg=COLORS['background'],
            fg=COLORS['text'],
            font=('Segoe UI', 10)
        )

if __name__ == "__main__":
    import os
    root = create_modern_gui()
    root.mainloop()