import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import React from "react";
import { LookdownControlProps } from "../types/typings";
import LookdownControlClassic from "./LookdownControlClassic";
import LookdownControlNewLook from "./LookdownControlNewLook";

const queryClient = new QueryClient();

export default function LookdownControl(props: LookdownControlProps) {
  const isNewLook = !!props.fluentDesign;
  return (
    <QueryClientProvider client={queryClient}>
      {isNewLook ? <LookdownControlNewLook {...props} /> : <LookdownControlClassic {...props} />}
    </QueryClientProvider>
  );
}
