/** Process List (공정편성 LIST) of a model: one row per process with its people, standard time, machine and jigs. */
export default function ProcessListTab(_props: { modelName: string; onError: (e: unknown) => void }) {
  return (
    <section className="card pi-sec">
      <div className="pi-sec-h"><h2>Process List</h2></div>
    </section>
  );
}
