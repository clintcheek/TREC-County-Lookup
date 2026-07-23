# Texas Broker County Resolver v7.7.1

## Search sequence update

Each brokerage now begins with a simple, human-style lookup sequence:

1. Brokerage name + address/phone.
2. Brokerage name + responsible broker name + address/phone.
3. Responsible broker name + “real estate agent” or “real estate broker” + address/phone.
4. License-based corroboration only after the simple searches.

## County confirmation

Whenever an address candidate is geocoded, the resolver also sends the plain-language query:

`what county is this address in "[address]"`

The U.S. Census geocoder remains the primary structured source. The search answer confirms the Census county when it agrees and can provide a fallback county when Census returns an address without a county. Conflicting search answers do not replace the Census county.
