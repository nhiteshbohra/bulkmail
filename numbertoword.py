from num2words import num2words

start = 1
end = 500

with open("numbers.txt", "w", encoding="utf-8") as f:
    for number in range(start, end + 1):
        f.write(f"{'escalation ' + num2words(number) + ' :- '}{' How Is This My Mistake? – Request for Written Justification from Amazon'}\n")

print("Done! Saved to numbers.txt")