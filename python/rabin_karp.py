class rolling_hash:
    def __init__(self, text, patternSize):
        self.text = text
        self.patternSize = patternSize
        # Увеличиваем base для поддержки большего алфавита (русский + английский)
        self.base = 256  # Используем больший base для Unicode символов
        self.window_start = 0
        self.window_end = 0
        self.mod = 5807
        self.hash = self.get_hash(text, patternSize)

    def get_hash(self, text, patternSize):
        hash_value = 0
        for i in range(0, patternSize):
            # Используем ord() напрямую для поддержки Unicode (включая русский язык)
            char_value = ord(self.text[i])
            hash_value = (
                hash_value + char_value * (self.base**(patternSize - i - 1))) % self.mod

        self.window_start = 0
        self.window_end = patternSize

        return hash_value

    def next_window(self):
        if self.window_end <= len(self.text) - 1:
            # Используем ord() напрямую для поддержки Unicode (включая русский язык)
            self.hash -= ord(self.text[self.window_start]) * self.base**(self.patternSize-1)

            self.hash *= self.base
            self.hash += ord(self.text[self.window_end])
            self.hash %= self.mod
            self.window_start += 1
            self.window_end += 1
            return True
        return False

    def current_window_text(self):
        return self.text[self.window_start:self.window_end]


def checker(text, pattern):
    if text == "" or pattern == "":
        return None
    if len(pattern) > len(pattern):
        return None

    text_rolling = rolling_hash(text.lower(), len(pattern))
    pattern_rolling = rolling_hash(pattern.lower(), len(pattern))

    for _ in range(len(text)-len(pattern)+1):
        print(pattern_rolling.hash, text_rolling.hash)
        if text_rolling.hash == pattern_rolling.hash:
            return "Found"
        text_rolling.next_window()
    return "Not Found"


if __name__ == "__main__":
    print(checker("ABDCCEAGmsslslsosspps", "agkalallaa"))
