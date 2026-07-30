from num2words import num2words

output_file = "ordinal_numbers.txt"

with open(output_file, "w", encoding="utf-8") as f:
    for i in range(1002, 2001):
        f.write(num2words(i, to="ordinal").capitalize() + "\n")

print(f"Saved to {output_file}")