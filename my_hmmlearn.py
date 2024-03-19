from hmmlearn import hmm
import numpy as np

# Define observation sequences
X = np.array([[0, 1, 0, 1], [1, 1, 0, 0], [0, 0, 0, 1]])

# Create and train the Hidden Markov Model
model = hmm.MultinomialHMM(n_components=2, n_iter=100)
model.fit(X)

# Decode the most likely sequence of hidden states for a given sequence
log_prob, hidden_states = model.decode(X)
print("Most likely sequence of hidden states:", hidden_states)

# Generate a new sequence of observations
generated_sequence, _ = model.sample(n_samples=5)
print("Generated sequence:", generated_sequence)
