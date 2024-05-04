import docx
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import string
import nltk
import csv

# Step 1: Extract text from the document
doc = docx.Document('EJ Marketting New.docx')
text = ' '.join([paragraph.text for paragraph in doc.paragraphs])

# Step 2: Preprocess the text (e.g., remove punctuation and stop words)
# Remove punctuation
punctuation = string.punctuation
text_no_punctuation = ''.join([char for char in text if char not in punctuation])

# Remove stop words
stopwords = nltk.corpus.stopwords.words('english')
text_no_stopwords = ' '.join([word for word in text_no_punctuation.split() if word.lower() not in stopwords])

# Remove other words
otherwords = ["for", "and", "my", "a", "of", "the", "on", "to", "role", "Name:", "project", "as", "from", "across", "&", "like", "that", "more",
              "Project", "Data", "data"]
final_text = ' '.join([word for word in text_no_stopwords.split() if word.lower() not in otherwords])

# Step 3: Calculate word frequencies
word_counts = {}

for word in final_text.split():
    if word not in word_counts:
        word_counts[word] = 0
    word_counts[word] += 1

print(word_counts)

# Open a CSV file for writing
with open('data.csv', 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    for key, value in word_counts.items():
        if isinstance(value, list) or isinstance(value, dict):
            for item in value:
                writer.writerow([key, item])
        else:
            writer.writerow([key, value])


# Step 4: Generate the word cloud
wordcloud = WordCloud(width=800, height=400).generate_from_frequencies(word_counts)

# Display the generated word cloud using matplotlib
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud, interpolation='bilinear')
plt.axis('off')
plt.show()