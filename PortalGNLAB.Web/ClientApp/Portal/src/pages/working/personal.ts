// Personal under Working: each user's own plans, kept as the issues of a model named after the user (the server lets only
// that user open it; WorkingEndpoints).

const PREFIX = "~personal:";

/** The model name a user's personal plans are kept under. */
export const personalKey = (user: string) => PREFIX + user.trim();
