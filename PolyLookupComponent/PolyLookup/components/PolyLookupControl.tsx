import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import React from "react";
import { PolyLookupProps } from "../types/typings";
import PolyLookupControlClassic from "./PolyLookupControlClassic";
import PolyLookupControlNewLook from "./PolyLookupControlNewLook";

const queryClient = new QueryClient();

export default function PolyLookupControl(props: PolyLookupProps) {
  const isNewLook = !!props.fluentDesign;
  return (
    <QueryClientProvider client={queryClient}>
      {isNewLook ? <PolyLookupControlNewLook {...props} /> : <PolyLookupControlClassic {...props} />}
    </QueryClientProvider>
  );
}
