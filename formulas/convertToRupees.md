The formula dynamically converts a numeric value into a more readable format using common PKR numbering system units (e.g., Lakh, Crore, Arab). It shortens large numbers into units based on their scale, making them easier to read and understand.

**Breakdown of Units:**

Lakh (1,00,000): Used for values in the range of 1,00,000 to 9,99,999.

Crore (1,00,00,000): Used for values in the range of 1,00,00,000 to 99,99,99,999.

Arab (1,00,00,00,000): Used for 1,00,00,00,000 and above.

```
=LET(
  value, C2,
  IFS(
    value >= 1000000000, ROUND(value/1000000000, 2) & " Arab",
    value >= 10000000, ROUND(value/10000000, 2) & " Crore",
    value >= 100000, ROUND(value/100000, 2) & " Lakh",
    TRUE, value
  )
)
```
