from dataclasses import dataclass


@dataclass(frozen=True)
class Tool:
    slug: str
    name: str
    description: str
    href: str
    protected: bool = False
    available: bool = True


TOOLS: list[Tool] = [
    Tool(
        slug="cc-analyzer",
        name="CC Analyzer Plus",
        description=(
            "Reconcile a Security Bank credit card PDF bill against your "
            "Money Manager data. Flags missing, duplicate, and inaccurate "
            "entries before pushing to Google Sheets."
        ),
        href="/tools/cc-analyzer",
        protected=True,
        available=True,
    ),
    Tool(
        slug="money-manager",
        name="Analyze Money Manager",
        description=(
            "Visualise spending patterns from a Money Manager xlsx export — "
            "top merchants, food/coffee/grocery breakdowns."
        ),
        href="/tools/money-manager",
        protected=False,
        available=False,
    ),
    Tool(
        slug="weather",
        name="Is Weather Good Here?",
        description="Check weather conditions in a given location.",
        href="/tools/weather",
        protected=False,
        available=False,
    ),
    Tool(
        slug="coordinates",
        name="Get Coordinates",
        description="Get latitude and longitude of a given location.",
        href="/tools/coordinates",
        protected=False,
        available=False,
    ),
]
