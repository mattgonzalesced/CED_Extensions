options = (["FLA","MOCP", "OCPD", 3, "FLA"])

print(options.index("FLA"))

if "MOCP" in options:
    print(options.count("FLA"))
    print("FLA occurs", (options.count("FLA")), "times")
    print("at", options.index("FLA"))

else:
    print("MOCP not here")


def say_hello():
    print("hello world")

say_hello()