# french-llm-evaluation
Master's research project on French grammatical error correction using LLMs, synthetic data augmentation, Stanza, token-level alignment and human-in-the-loop evaluation.

## Research Reproducibility

The original evaluation pipeline was successfully
reproduced locally on macOS using Python 3.12.

The reproduction was conducted on the complete research
corpus of 107 French sentence pairs and 428 original
LLM outputs across four evaluated models.

The workflow includes:

- French morphosyntactic analysis with Stanza
- Token-level alignment using Levenshtein distance
- Edit extraction and comparison
- Normalized exact-match accuracy
- Edit-level Precision, Recall and F1
- Performance analysis by linguistic error category
- Preservation of existing human annotations
- Automated Excel evaluation report generation

The reproduced global metrics and category-level
accuracy scores are consistent with the results
reported in the final master's thesis.

The original human annotations were preserved,
including 30 acceptable alternative corrections.

The reproduction uses the original model outputs
without requiring new external API calls.
