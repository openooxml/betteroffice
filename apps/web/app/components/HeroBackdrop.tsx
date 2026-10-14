"use client";

import { useEffect, useState } from "react";
import { HeroField } from "./HeroField";
import { HeroStream } from "./HeroStream";

// what fills the hero above the copy; ?field=stream|perspective|empty picks one while they are compared
const VARIANTS = ["stream", "perspective", "empty"] as const;
type Variant = (typeof VARIANTS)[number];

export function HeroBackdrop() {
  const [variant, setVariant] = useState<Variant>("stream");
  useEffect(() => {
    const asked = new URLSearchParams(window.location.search).get("field");
    if (VARIANTS.includes(asked as Variant)) setVariant(asked as Variant);
  }, []);
  if (variant === "empty") return null;
  return variant === "perspective" ? <HeroField /> : <HeroStream />;
}
