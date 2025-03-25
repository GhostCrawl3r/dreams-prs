import axios, { AxiosResponse } from "axios";
import base64 from "base-64";
import * as XLSX from "xlsx";
import * as dotenv from "dotenv";

dotenv.config();

const USERNAME = process.env.BITBUCKET_USERNAME as string;
const APP_PASSWORD = process.env.BITBUCKET_APP_PASSWORD as string;
const WORKSPACE = process.env.BITBUCKET_WORKSPACE as string;
const REPO_SLUG = "dreams"

if (!USERNAME || !APP_PASSWORD || !WORKSPACE || !REPO_SLUG) {
    throw new Error("Missing required environment variables in .env file");
}

const authHeader = "Basic " + base64.encode(`${USERNAME}:${APP_PASSWORD}`);

const ninetyDaysAgo = new Date();
ninetyDaysAgo.setDate(ninetyDaysAgo.getDate() - 90);
const sinceDate = ninetyDaysAgo.toISOString();

type PRState = "OPEN" | "MERGED";

const formatDate = (isoString: string) => {
    return new Date(isoString).toISOString().replace("T", " ").split(".")[0];
};

async function getContributorsForPR(prId: number) {
    const url = `https://api.bitbucket.org/2.0/repositories/${WORKSPACE}/${REPO_SLUG}/pullrequests/${prId}/commits`;

    try {
        let contributors = new Set<string>();
        let nextUrl: string | null = url;

        while (nextUrl) {
            const response: any = await axios.get(nextUrl, {
                headers: { Authorization: authHeader }
            });

            response.data.values.forEach((commit: any) => {
                if (commit.author?.user?.display_name) {
                    contributors.add(commit.author.user.display_name);
                } else if (commit.author?.raw) {
                    contributors.add(commit.author.raw.split("<")[0].trim());
                }
            });

            nextUrl = response.data.next || null;
        }

        return Array.from(contributors);
    } catch (error: any) {
        console.error(`Error fetching contributors for PR ${prId}:`, error.response?.data || error.message);
        return [];
    }
}

async function fetchPRs(state: PRState) {
    const url = `https://api.bitbucket.org/2.0/repositories/${WORKSPACE}/${REPO_SLUG}/pullrequests?state=${state}&pagelen=50`;

    try {
        let allPRs: any[] = [];
        let nextUrl: string | null = url;

        while (nextUrl) {
            const response: AxiosResponse = await axios.get(nextUrl, {
                headers: { Authorization: authHeader }
            });

            const filteredPRs = await Promise.all(
                response.data.values
                    .filter((pr: any) => new Date(pr.created_on) >= new Date(sinceDate))
                    .map(async (pr: any) => {
                        const contributors = await getContributorsForPR(pr.id);
                        return {
                            ID: pr.id,
                            Title: pr.title,
                            State: pr.state,
                            Developer: pr.author.display_name,
                            Created_On: formatDate(pr.created_on),
                            Updated_On: formatDate(pr.updated_on),
                            Source_Branch: pr.source?.branch?.name || "Unknown",
                            Destination_Branch: pr.destination?.branch?.name || "Unknown",
                            PR_Link: pr.links?.html?.href || "N/A",
                            Contributors: [...new Set([pr.author.display_name, ...contributors])].join(", "),
                            Comment_Count: pr.comments_count || 0,
                            Closed_By: pr.closed_by?.display_name || "N/A",
                        };
                    })
            );

            allPRs = [...allPRs, ...filteredPRs];
            nextUrl = response.data.next || null;
        }

        console.log(`Total ${state} PRs in the last 90 days: ${allPRs.length}`);
        return allPRs;
    } catch (error: any) {
        console.error(`Error fetching ${state} PRs:`, error.response?.data || error.message);
        return [];
    }
}

async function exportToExcel() {
    const [openPRs, mergedPRs] = await Promise.all([
        fetchPRs("OPEN"),
        fetchPRs("MERGED")
    ]);

    const allPRs = [...openPRs, ...mergedPRs];

    if (allPRs.length === 0) {
        console.log("No PRs found. No Excel file generated.");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(allPRs);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pull Requests");

    const filePath = `${REPO_SLUG}_prs.xlsx`;
    XLSX.writeFile(wb, filePath);
    console.log(`Excel file successfully saved as ${filePath}`);
}

exportToExcel().catch(error => console.log('error', error));
