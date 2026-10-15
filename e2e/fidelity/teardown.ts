import { writeReport } from './report'

/** Writes fidelity-report/index.html after every run that produced results. */
export default async function teardown() {
  writeReport()
}
