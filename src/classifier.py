def classify(text_lines, rules):
    content = " ".join(text_lines)
    scores = {}

    for cat, cfg in rules['categories'].items():
        score = sum(1 for kw in cfg['keywords'] if kw in content)
        if score:
            scores[cat] = score * cfg['priority']

    if not scores:
        return None, 0

    best = max(scores, key=scores.get)
    return best, scores[best]